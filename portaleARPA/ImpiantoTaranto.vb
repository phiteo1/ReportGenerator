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

Public Class ImpiantoTaranto
    Implements IImpianto

    Private formInstance As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
    Dim connectionString As String
    Dim connectionStringCTE As String
    Private _chimneyList As New List(Of Camino)
    Private actualState As Byte
    Private cteConfiguration As String
    Private cteInvertedConfiguration As String
    Private O2RefDict As Dictionary(Of String, Integer)
    Private hiddenColumns As New List(Of String)()
    Private culture As System.Globalization.CultureInfo
    Private ret As Int32
    Private ret2 As Int32
    Dim datanh3 As String
    Dim mesenh3 As Integer
    Dim hnf, htran, vleCo, vleNox As String


    Enum State                  'State Machine of the downloading process
        DataLoading = 1
        TableLoading = 2
        SheetLoading = 3
        FinishedReport = 4
        Finished = 5
    End Enum

    Public Sub New()

        AddChimneyToList()
        culture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        culture.NumberFormat.NumberGroupSeparator = ""
        cteConfiguration = ""
        cteInvertedConfiguration = ""
        O2RefDict = Nothing
        connectionString = ConfigurationManager.ConnectionStrings("AQMSDBCONN").ConnectionString
        connectionStringCTE = ConfigurationManager.ConnectionStrings("AQMSDBCONNCTE").ConnectionString
        datanh3 = ConfigurationManager.AppSettings("datanh3")
        mesenh3 = ConfigurationManager.AppSettings("mesenh3")

    End Sub


    Private Sub AddChimneyToList() Implements IImpianto.AddChimneyToList

        _chimneyList.Add(New Camino("E1", 1))
        _chimneyList.Add(New Camino("E2", 2))
        _chimneyList.Add(New Camino("E3", 1))
        _chimneyList.Add(New Camino("E4", 3))
        _chimneyList.Add(New Camino("E7", 4))
        _chimneyList.Add(New Camino("E8", 5))
        _chimneyList.Add(New Camino("E9", 6))
        _chimneyList.Add(New Camino("E10", 7))
        _chimneyList.Add(New Camino("Flussi di massa", 8, 0))
        _chimneyList.Add(New Camino("Bolla di raffineria", 8, 1))

        

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

        'Console.WriteLine(": " & Form1.startDate & Form1.endDate & Form1.section & Form1.reportType)

        Dim exePath As String = Application.StartupPath                                                                                                                     ' Get the 2 layer up directory
        Dim grandParentPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim chimneyName As String = MySharedMethod.GetChimneyName(Convert.ToInt16(Form1.section.ToString()))
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

        If Form1.isCte Then

            If O2RefDict Is Nothing Then
                O2RefDict = New Dictionary(Of String, Integer)
                O2RefDict.Add("cogenerativo", 15)
                O2RefDict.Add("caldaia", 3)
            End If

            Dim invertedIndex As Byte
            invertedIndex = If(GetComboBoxSelectedIndex(formInstance, formInstance.ComboBox3) = 0, 1, 0)
            cteConfiguration = LCase(GetComboBoxSelectedItem(formInstance, formInstance.ComboBox3))
            cteInvertedConfiguration = LCase(GetComboBoxSelectedItemFromIndex(formInstance, formInstance.ComboBox3, invertedIndex))
        End If

        If Form1.isCte Then
            reportPath = Path.Combine(grandParentPath, "report", "E3")
        End If

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
            If Form1.section = 8 Then
                If Form1.bolla = 0 Then
                    dataTable1 = GetDataFlussi(barProgress, startDate, endDate, Form1.section, Form1.reportType, 1)                                                                            'Get the data from the database and assign to first data table structure. The function is runned in an other trhead in order to allow the GUI to refresh properly
                    dataTable2 = GetDataFlussi(barProgress, startDate, Form1.endDate, Form1.section, Form1.reportType, 2)                                                                           'Get the data from the database and assign to second data table structure
                    preRenderFirstTable(Form1.section)
                ElseIf Form1.bolla = 1 Then
                    dataTable1 = GetFirstBollaTable(barProgress, Form1.startDate, Form1.endDate, Form1.section, Form1.reportType)                                                                          'Get the data from the database and assign to first data table structure. The function is runned in an other trhead in order to allow the GUI to refresh properly
                    dataTable2 = GetSecondBollaTable(barProgress, Form1.startDate, Form1.endDate, Form1.section, Form1.reportType)                                                                        'Get the data from the database and assign to second data table structure
                Else
                    MessageBox.Show("Errore nella scelta della configurazione del camino.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    EnableFormSafe(formInstance)
                    Return
                End If

            Else
                If Form1.isCte = False Then
                    dataTable1 = GetFirstCaminiTable(barProgress, startDate, endDate, Form1.section, Form1.reportType)
                    dataTable2 = GetSecondCaminiTable(barProgress, startDate, endDate, Form1.section, Form1.reportType)
                    preRenderFirstTable(Form1.section)
                Else
                    dataTable1 = GetFirstCTETable(barProgress, startDate, endDate, Form1.section, Form1.reportType)
                    dataTable2 = GetSecondCTETable(barProgress, startDate, endDate, Form1.section, Form1.reportType)
                End If
            End If

            If dataTable1 Is Nothing Then
                MessageBox.Show("Errore nell'acquisizione dei dati, consultare il file di log per i dettagli.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                EnableFormSafe(formInstance)
                Return
            Else
                UpdateDgvDataSource(dataTable1, Form1.dgv)                                                                                                                                    'Bind the data to the first DataGridView
            End If

            If dataTable2 Is Nothing Then
                MessageBox.Show("Errore nell'acquisizione dei dati, consultare il file di log per i dettagli.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                EnableFormSafe(formInstance)
                Return
            Else
                UpdateDgvDataSource(dataTable2, Form1.dgv2)                                                                                                                                   'Bind the data to the second DataGridView
            End If

            If Form1.bolla = 0 Then
                downloadReportFlussi(StatusProgress, startDate, endDate, reportPath)                                                                                                    'Download the reports of the selected years(months).

            ElseIf Form1.bolla = 1 Then

                downloadReportBolla(StatusProgress, startDate, endDate, reportPath)

            Else


                If Form1.reportType = 0 Then
                    If Form1.isCte = False Then
                        downloadYearlyReportCamini(StatusProgress, startDate, endDate, reportPath)
                    Else
                        downloadYearlyReportCTE(StatusProgress, startDate, endDate, reportPath)
                    End If

                ElseIf Form1.reportType = 1 Then

                    If (Form1.isCte = False) Then
                        downloadMonthlyReportCamini(StatusProgress, startDate, endDate, reportPath)
                    Else
                        downloadMonthlyReportCTE(StatusProgress, startDate, endDate, reportPath)
                    End If

                End If

            End If

            Dim deltaTime As String
            If (Form1.reportType = 0) Then
                deltaTime = "yyyy"                                                                                                                                                      'Add one year or one month according to the report type choosed
            Else
                deltaTime = "m"
            End If
            startDate = DateAdd(deltaTime, 1, startDate)

        End While

    End Sub


    Private Function GetDataFlussi(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32, whatTable As Byte) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim commandCTE As System.Data.SqlClient.SqlCommand
        Dim connection As New SqlConnection(connectionString)
        Dim connectionCTE As New SqlConnection(connectionStringCTE)
        Dim queryNumber As Integer = 0
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim dataType As String = " AND TIPO_DATO IS NOT NULL ORDER BY INS_ORDER"
        Dim methodName As String = GetCurrentMethod()

        Try
            ' Tenta di aprire la connessione
            connection.Open()
            connectionCTE.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
            Return dt
        End Try

        Select Case Form1.reportType
            Case 0
                datanh3 = "01/01/2020"
        End Select

        progress.Report(State.DataLoading)
        dt.Columns.Add(New Data.DataColumn("IDX_REPORT", GetType(Double)))
        dt.Columns.Add(New Data.DataColumn("INS_ORDER", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("ORA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E2Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E3Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E4Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E7Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E8Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E9Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_COV", GetType(String)))


        dt.Columns.Add(New Data.DataColumn("E9Q_NH3", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E10Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_COV", GetType(String)))
        If whatTable = 1 Then
            dt.Columns.Add(New Data.DataColumn("NOX_SOMMA", GetType(String)))
            dt.Columns.Add(New Data.DataColumn("SO2_SOMMA", GetType(String)))
            dt.Columns.Add(New Data.DataColumn("POLVERI_SOMMA", GetType(String)))


            dt.Columns.Add(New Data.DataColumn("CO_SOMMA", GetType(String)))


            dt.Columns.Add(New Data.DataColumn("COV_SOMMA", GetType(String)))
            dt.Columns.Add(New Data.DataColumn("NH3_SOMMA", GetType(String)))

            dt.Columns.Add(New Data.DataColumn("NOX57_SOMMA", GetType(String)))
            Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_MASSICI_CAMINI", connection)
            testCMD.CommandTimeout = 18000
            testCMD.CommandType = Data.CommandType.StoredProcedure
            testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
            testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@idsez").Value = section
            testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
            testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@data").Value = startTime
            testCMD.Parameters.Add("@TIPO_ESTRAZIONE", Data.SqlDbType.Int, 11)
            testCMD.Parameters("@TIPO_ESTRAZIONE").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@TIPO_ESTRAZIONE").Value = Form1.reportType
            testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
            testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
            Try
                testCMD.ExecuteScalar()
            Catch ex As Exception
                Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
                dt = Nothing
                Return dt
            End Try

            ret = testCMD.Parameters("@retval").Value
            testCMD = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_MASSICI_CAMINI", connectionCTE)
            testCMD.CommandTimeout = 18000
            testCMD.CommandType = Data.CommandType.StoredProcedure
            testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
            testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@idsez").Value = section
            testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
            testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@data").Value = startTime
            testCMD.Parameters.Add("@TIPO_ESTRAZIONE", Data.SqlDbType.Int, 11)
            testCMD.Parameters("@TIPO_ESTRAZIONE").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@TIPO_ESTRAZIONE").Value = Form1.reportType
            testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
            testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
            testCMD.Parameters("@idsez").Value = 1
            Try
                testCMD.ExecuteScalar()
            Catch ex As Exception
                Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
                dt = Nothing
                Return dt
            End Try

            ret2 = testCMD.Parameters("@retval").Value
            queryNumber += 3
            progress.Report(queryNumber * progressStep)

            dataType = " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"

        End If

        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim logStatement As String = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & dataType
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)
        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt

        End Try

        Dim reader2 As System.Data.SqlClient.SqlDataReader
        logStatement = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret2.ToString() & dataType
        commandCTE = New System.Data.SqlClient.SqlCommand(logStatement, connectionCTE)
        Try
            reader2 = commandCTE.ExecuteReader()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try

        Dim dr As Data.DataRow = dt.NewRow()

        If (reader.HasRows) Then
            While reader.Read()
                Try
                    reader2.Read()
                    dr("IDX_REPORT") = reader("IDX_REPORT")
                    dr("INS_ORDER") = String.Format("{0:n0}", reader("INS_ORDER"))
                    dr("ORA") = reader("ORA") 'String.Format("{0:n2}", reader("NOX"))

                    dr("E1Q_NOX") = reader("E1Q_NOX")
                    dr("E1Q_SO2") = reader("E1Q_SO2")
                    dr("E1Q_POLVERI") = reader("E1Q_POLVERI")
                    dr("E1Q_CO") = reader("E1Q_CO")
                    dr("E1Q_COV") = reader("E1Q_COV")

                    dr("E2Q_NOX") = reader("E2Q_NOX")
                    dr("E2Q_SO2") = reader("E2Q_SO2")
                    dr("E2Q_POLVERI") = reader("E2Q_POLVERI")
                    dr("E2Q_CO") = reader("E2Q_CO")
                    dr("E2Q_COV") = reader("E2Q_COV")

                    Try
                        dr("E3Q_NOX") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_NOX"), culture.NumberFormat))
                    Catch e As FormatException         ''il dato non è un double
                        dr("E3Q_NOX") = reader2("E1Q_NOX")
                    Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException ''non c'è il dato per E3
                        dr("E3Q_NOX") = "--"
                    End Try

                    Try
                        dr("E3Q_SO2") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_SO2"), culture.NumberFormat))
                    Catch e As FormatException
                        dr("E3Q_SO2") = reader2("E1Q_SO2")
                    Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException     ''non c'è il dato per E3
                        dr("E3Q_SO2") = "--"
                    End Try

                    Try
                        dr("E3Q_POLVERI") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_POLVERI"), culture.NumberFormat))
                    Catch e As FormatException
                        dr("E3Q_POLVERI") = reader2("E1Q_POLVERI")
                    Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException     ''non c'è il dato per E3
                        dr("E3Q_POLVERI") = "--"
                    End Try

                    Try
                        dr("E3Q_CO") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_CO"), culture.NumberFormat))
                    Catch e As FormatException         ''il dato non è un double
                        dr("E3Q_CO") = reader2("E1Q_CO")
                    Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException ''non c'è il dato per E3
                        dr("E3Q_CO") = "--"
                    End Try

                    Try
                        dr("E3Q_COV") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_COV"), culture.NumberFormat))
                    Catch e As FormatException         ''il dato non è un double
                        dr("E3Q_COV") = reader2("E1Q_COV")
                    Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException ''non c'è il dato per E3
                        dr("E3Q_COV") = "--"
                    End Try

                    dr("E4Q_NOX") = reader("E4Q_NOX")
                    dr("E4Q_SO2") = reader("E4Q_SO2")
                    dr("E4Q_POLVERI") = reader("E4Q_POLVERI")
                    dr("E4Q_CO") = reader("E4Q_CO")
                    dr("E4Q_COV") = reader("E4Q_COV")

                    dr("E7Q_NOX") = reader("E7Q_NOX")
                    dr("E7Q_SO2") = reader("E7Q_SO2")
                    dr("E7Q_POLVERI") = reader("E7Q_POLVERI")
                    dr("E7Q_CO") = reader("E7Q_CO")
                    dr("E7Q_COV") = reader("E7Q_COV")

                    dr("E8Q_NOX") = reader("E8Q_NOX")
                    dr("E8Q_SO2") = reader("E8Q_SO2")
                    dr("E8Q_POLVERI") = reader("E8Q_POLVERI")
                    dr("E8Q_CO") = reader("E8Q_CO")
                    dr("E8Q_COV") = reader("E8Q_COV")

                    dr("E9Q_NOX") = reader("E9Q_NOX")
                    dr("E9Q_SO2") = reader("E9Q_SO2")
                    dr("E9Q_POLVERI") = reader("E9Q_POLVERI")
                    dr("E9Q_CO") = reader("E9Q_CO")
                    dr("E9Q_COV") = reader("E9Q_COV")
                    If (reader("E9Q_NH3") IsNot DBNull.Value) Then
                        dr("E9Q_NH3") = reader("E9Q_NH3")
                    Else
                        dr("E9Q_NH3") = "0"
                    End If

                    dr("E10Q_NOX") = reader("E10Q_NOX")
                    dr("E10Q_SO2") = reader("E10Q_SO2")
                    dr("E10Q_POLVERI") = reader("E10Q_POLVERI")
                    dr("E10Q_CO") = reader("E10Q_CO")
                    dr("E10Q_COV") = reader("E10Q_COV")

                    If whatTable = 1 Then
                        dr("NOX_SOMMA") = reader("NOX_SOMMA")
                        dr("SO2_SOMMA") = reader("SO2_SOMMA")
                        dr("POLVERI_SOMMA") = reader("POLVERI_SOMMA")
                        dr("CO_SOMMA") = reader("CO_SOMMA")
                        dr("COV_SOMMA") = reader("COV_SOMMA")
                        dr("NH3_SOMMA") = reader("NH3_SOMMA")
                        dr("NOX57_SOMMA") = reader("NOX57_SOMMA")
                    Else
                        If (reader("TIPO_DATO").ToString().Contains("DISP")) Then
                            For i As Integer = 3 To dr.Table.Columns.Count - 1 Step 1
                                dr(i) = String.Format(culture, "{0:P2}", Double.Parse(dr(i), culture.NumberFormat))
                            Next
                        ElseIf (reader("TIPO_DATO").ToString().Contains("AVG") Or reader("TIPO_DATO").ToString().Contains("Totale")) Then
                            For i As Integer = 3 To dr.Table.Columns.Count - 1 Step 1
                                dr(i) = String.Format(culture, "{0:n2}", Double.Parse(dr(i), culture.NumberFormat))
                            Next
                        ElseIf (reader("TIPO_DATO").ToString().Contains("N.F.") Or reader("TIPO_DATO").ToString().Contains("VALIDITA")) Then
                            For i As Integer = 3 To dr.Table.Columns.Count - 1 Step 1
                                dr(i) = String.Format(culture, "{0:n0}", Double.Parse(dr(i), culture.NumberFormat))
                            Next
                        End If
                    End If

                    dt.Rows.Add(dr)
                    dr = dt.NewRow()

                Catch ex As Exception
                    Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                    Continue While
                End Try

            End While

            queryNumber += 1
            progress.Report(queryNumber * progressStep)

        End If





        If (startTime < Date.Parse(datanh3)) Then
            hiddenColumns.Add("E9Q_NH3")
            If whatTable = 1 Then
                hiddenColumns.Add("NH3_SOMMA")
                hiddenColumns.Add("NOX57_SOMMA")
            End If
        End If

        connection.Close()
        connectionCTE.Close()

        Return dt
    End Function

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
        Dim retLong As Long

        If Form1.reportType = 0 Then                                                      ' It was needed thanks to the genius who wrote the logics in the portal :))
            type = 3
        ElseIf Form1.reportType = 1 Then
            type = 2
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



        dt.Columns.Add(New Data.DataColumn("INTESTAZIONE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("COV", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NH3", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("H2O", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2RIF", GetType(String)))

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_REPORT_WEB", connection)
        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = section

        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = Format("{0:dd/MM/yyyy}", startTime)


        testCMD.Parameters.Add("@aia", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@aia").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@aia").Value = Form1.aia


        testCMD.Parameters.Add("@tipoestrazione", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@tipoestrazione").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@tipoestrazione").Value = type

        testCMD.Parameters.Add("@retval", Data.SqlDbType.BigInt, 8)
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

        retLong = testCMD.Parameters("@retval").Value

        hnf = testCMD.Parameters("@HNF").Value.ToString()
        htran = testCMD.Parameters("@H_TRANS").Value.ToString()

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Dim log_statement As String = "SELECT * FROM [ARPA_REPORT_WEB] WHERE IDX_REPORT = " & retLong.ToString() & dataType
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

                If type = 2 Then
                    If (dr("INTESTAZIONE") = "VLE GIC [mg/Nm3]" And section = 3 Or section = 4 Or section = 7 And Form1.aia = 0) Then
                        dr = dt.NewRow()
                    End If
                End If

                If (IsNumeric(reader("CO"))) Then
                    If (Convert.ToDouble(reader("CO")) >= 0) Then
                        dr("CO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO"))
                    Else
                        dr("CO") = "N.A."
                    End If
                Else
                    dr("CO") = "N.A."
                End If


                If (IsNumeric(reader("NOX"))) Then
                    If (Convert.ToDouble(reader("NOX")) >= 0) Then
                        dr("NOX") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX"))
                    Else
                        dr("NOX") = "N.A."
                    End If
                Else
                    dr("NOX") = "N.A."
                End If

                'Inserisce N.A. quando i vle degli inquinanti non sono presenti

                If (IsNumeric(reader("SO2"))) Then
                    If (Convert.ToDouble(reader("SO2")) >= 0) Then
                        dr("SO2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2"))

                    Else
                        dr("SO2") = "N.A."
                    End If
                Else
                    dr("SO2") = "N.A."
                End If



                If (IsNumeric(reader("POLVERI"))) Then
                    If (Convert.ToDouble(reader("POLVERI")) > 0) Then
                        dr("POLVERI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POLVERI"))
                    Else
                        dr("POLVERI") = "N.A."
                    End If
                Else
                    dr("POLVERI") = "N.A."
                End If


                If (IsNumeric(reader("COV"))) Then
                    If (Convert.ToDouble(reader("COV")) > 0) Then
                        dr("COV") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COV"))
                    Else
                        dr("COV") = "N.A."
                    End If
                Else
                    dr("COV") = "N.A."
                End If
                '      If ((SelectedDate.Value > datanh3) And (String.Equals(Sezione.Text, "6"))) Then
                If (IsNumeric(reader("NH3"))) Then
                    If (Convert.ToDouble(reader("NH3")) > 0) Then
                        dr("NH3") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NH3"))
                    Else
                        dr("NH3") = "N.A."
                    End If
                Else
                    dr("NH3") = "N.A."
                End If



                'If (String.Equals(Sezione.Text, "6")) Then

                'End If
                dr("O2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("O2"))
                dr("QFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("QFUMI"))
                dr("TFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("TFUMI"))
                dr("PFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PFUMI"))
                dr("H2O") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("H2O"))
                dr("O2RIF") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("O2RIF"))
                'dr("MWE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"),"{0:n2}", reader("MWE"))

                dt.Rows.Add(dr)
                dr = dt.NewRow()
            Catch ex As Exception
                Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                Continue While
            End Try

        End While

        If ((startTime < Date.Parse(datanh3)) Or (startTime >= Date.Parse(datanh3) And section <> 6)) Then
            hiddenColumns.Add("NH3")
        End If

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        connection.Close()

        Return dt
    End Function


    Private Function GetSecondCaminiTable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 3
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim methodName As String = GetCurrentMethod()
        Dim dataType As String = " ORDER BY INS_ORDER"
        Dim retLong As Long

        If Form1.reportType = 0 Then                                                      ' It was needed thanks to the genius who wrote the logics in the portal :))
            type = 3
        ElseIf Form1.reportType = 1 Then
            type = 2
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
        dt.Columns.Add(New Data.DataColumn("NOX_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_NOX", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("CO_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_CO", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("SO2_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_SO2", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("POLVERI_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POLVERI_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_POLVERI", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("COV_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("COV_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_COV", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_COV", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("NH3_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NH3_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_NH3", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2_MIS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2_RIF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("ORE_NF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("UFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_COV", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_NH3", GetType(String)))

        Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_MESE_ANNO_REPORT", connection)
        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = section 'Request.QueryString("Sezione").ToString())

        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = Format("{0:dd/MM/yyyy}", startTime) 'RepggCal.SelectedDate.ToString("dd/MM/yyyy HH:mm:ss")


        testCMD.Parameters.Add("@aia", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@aia").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@aia").Value = Form1.aia




        testCMD.Parameters.Add("@IS_MESE", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@IS_MESE").Direction = Data.ParameterDirection.Input
        If type = 3 Then
            testCMD.Parameters("@IS_MESE").Value = 0
        ElseIf type = 2 Then
            testCMD.Parameters("@IS_MESE").Value = 1
        End If
        testCMD.Parameters.Add("@retval", Data.SqlDbType.BigInt, 8)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_NOX", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_NOX").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_CO", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_CO").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_SO2", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_SO2").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_POLVERI", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_POLVERI").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_COV", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_COV").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_NH3", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_NH3").Direction = Data.ParameterDirection.Output

        Try
            testCMD.ExecuteScalar()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
            dt = Nothing
            Return dt
        End Try

        retLong = testCMD.Parameters("@retval").Value

        Dim log_statement As String = "SELECT * FROM [ARPA_WEB_MESE_ANNO_REPORT2] WHERE IDX_REPORT = " & retLong.ToString() & dataType
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
                dr("ORA") = reader("ORA") 'String.Format(CultureInfo.CreateSpecificCulture("it-IT"),"{0:n2}", reader("NOX"))
                dr("NOX_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX_IC"))
                dr("NOX_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX_VLE"))
                dr("DISP_NOX") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_NOX"))
                dr("IS_BOLD_NOX") = reader("IS_BOLD_NOX")
                dr("CO_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO_IC"))
                dr("CO_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO_VLE"))
                dr("DISP_CO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_CO"))
                'bold= superi segnalati in rosso nel mensile)
                dr("IS_BOLD_CO") = reader("IS_BOLD_CO")
                dr("SO2_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2_IC"))
                dr("SO2_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2_VLE"))
                dr("DISP_SO2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_SO2"))
                dr("IS_BOLD_SO2") = reader("IS_BOLD_SO2")
                dr("POLVERI_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POLVERI_IC"))
                dr("POLVERI_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POLVERI_VLE"))
                dr("DISP_POLVERI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_POLVERI"))
                dr("IS_BOLD_POLVERI") = reader("IS_BOLD_POLVERI")

                dr("COV_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COV_IC"))
                dr("COV_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COV_VLE"))
                dr("DISP_COV") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_COV"))

                'Inserimento colonna bold (per il supero dell'inquinante COV) I limiti VLE del COV sono presenti solo nella nuova AIA. (Mensile e annuale)
                dr("IS_BOLD_COV") = reader("IS_BOLD_COV")

                dr("O2_MIS") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("O2_MIS"))
                dr("O2_RIF") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("O2_RIF"))
                dr("TFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("TFUMI"))
                dr("PFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PFUMI"))
                dr("ORE_NF") = reader("ORE_NF")
                dr("QFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("QFUMI"))
                dr("UFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("UFUMI"))
                dr("PORTATA_CO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_CO"))
                dr("PORTATA_NOX") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_NOX"))
                dr("PORTATA_SO2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_SO2"))
                dr("PORTATA_POLVERI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_POLVERI"))
                dr("PORTATA_COV") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_COV"))
                dr("PORTATA_NH3") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_NH3"))
                dr("NH3_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NH3_IC"))
                dr("NH3_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NH3_VLE"))
                dr("DISP_NH3") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_NH3"))

                dt.Rows.Add(dr)
                dr = dt.NewRow()

            Catch ex As Exception
                Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                Continue While
            End Try

        End While

        If (startTime < Date.Parse(datanh3) Or startTime >= Date.Parse(datanh3) And section <> 6) Then
            hiddenColumns.Add("NH3_IC")
            hiddenColumns.Add("NH3_VLE")
            hiddenColumns.Add("DISP_NH3")
            hiddenColumns.Add("PORTATA_NH3")
        End If


        connection.Close()

        Return dt
    End Function




    Private Function GetFirstBollaTable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 0
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim methodName As String = GetCurrentMethod()

        Try
            ' Tenta di aprire la connessione
            connection.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
            Return dt
        End Try

        progress.Report(State.DataLoading)
        dt.Columns.Add(New Data.DataColumn("IDX_REPORT", GetType(Double)))
        dt.Columns.Add(New Data.DataColumn("INS_ORDER", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("ORA", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("SO2_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("CO_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("NOX_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("POL_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POL_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("COV_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("COV_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("FUMI_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("FUMI_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("TIPO_DATO", GetType(String)))

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Dim testCMD As Data.SqlClient.SqlCommand

        If startTime >= "01/01/2018" Then
            testCMD = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_CONCENTRAZIONI_CAMINI2", connection)
            testCMD.Parameters.Add("@aia", Data.SqlDbType.Int, 11)
            testCMD.Parameters("@aia").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@aia").Value = Form1.aia


        Else
            testCMD = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_CONCENTRAZIONI_CAMINI", connection)


        End If


        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.CommandTimeout = 18000
        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = startTime
        testCMD.Parameters.Add("@TIPO_ESTRAZIONE", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@TIPO_ESTRAZIONE").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@TIPO_ESTRAZIONE").Value = type
        testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
        Try
            testCMD.ExecuteScalar()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
            dt = Nothing
            Return dt
        End Try

        ret = testCMD.Parameters("@retval").Value

        queryNumber += 1
        progress.Report(queryNumber * progressStep)


        Dim logStatement As String = "SELECT * FROM [ARPA_WEB_CONCENTRAZIONI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & "  ORDER BY INS_ORDER"
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)

        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try


        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat
        nfi.NumberGroupSeparator = ""
        Dim dr As Data.DataRow = dt.NewRow()
        If (reader.HasRows) Then
            While reader.Read()
                Try
                    dr("IDX_REPORT") = reader("IDX_REPORT")
                    dr("INS_ORDER") = String.Format("{0:n0}", reader("INS_ORDER"))
                    If (type = 2) Then
                        dr("ORA") = String.Format("{0:HH.mm}", reader("ORA")) 'String.Format("{0:n2}", reader("NOX"))
                    ElseIf (type = 1) Then
                        dr("ORA") = String.Format("{0:dd}", reader("ORA"))
                    Else
                        dr("ORA") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:MMMM}", DateTime.Parse(reader("ORA").ToString()))
                    End If

                    ''Quando comincio a leggere i dati riassuntivi inserisco una riga vuota
                    If reader("TIPO_DATO").ToString().Contains("AVG") Then          ''Il valore di media non va nello specchietto subito sotto alla tabella principale, non in fondo alla tabella                    
                        Continue While
                    ElseIf (reader("TIPO_DATO").ToString().Contains("MAX")) Then    '' I valori di massimo e minimo vanno in fondo alla prima tabella (v. pre_render)
                        dr("ORA") = "MAX"
                    ElseIf (reader("TIPO_DATO").ToString().Contains("SUPERI")) Then
                        dr("ORA") = "N Sup. Medie Giorn."
                    ElseIf (reader("TIPO_DATO").ToString().Contains("MIN")) Then
                        dr("ORA") = "MIN"
                    ElseIf (reader("TIPO_DATO").ToString().Contains("VLE")) Then
                        dt.Rows.Add()
                        dr("ORA") = "VLE"
                    ElseIf (reader("TIPO_DATO").ToString() = "") Then               ''Se il valore è NULL vuol dire che sto leggendo un dato giornaliero, quindi non devo fare nulla di particolare

                    Else                                                            ''Negli altri casi ho raggiunto la fine della Tabella SQL con i dati di interesse e quindi esco senza scrivere altro
                        Exit While
                    End If

                    Dim availability As Double
                    dr("SO2_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2_SECCO"))
                    dr("SO2_AVAIL") = String.Format("{0:0.00}", reader("SO2_AVAIL")) & "%"
                    If (Double.TryParse(reader("SO2_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("SO2_SECCO") = dr("SO2_SECCO") + "(*)"
                        End If
                    End If


                    dr("CO_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO_SECCO"))
                    dr("CO_AVAIL") = String.Format("{0:0.00}", reader("CO_AVAIL")) & "%"
                    If (Double.TryParse(reader("CO_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("CO_SECCO") = dr("CO_SECCO") + "(*)"
                        End If
                    End If

                    dr("NOX_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX_SECCO"))
                    dr("NOX_AVAIL") = String.Format("{0:0.00}", reader("NOX_AVAIL")) & "%"
                    If (Double.TryParse(reader("NOX_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("NOX_SECCO") = dr("NOX_SECCO") + "(*)"
                        End If
                    End If

                    dr("POL_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POL_SECCO"))
                    dr("POL_AVAIL") = String.Format("{0:0.00}", reader("POL_AVAIL")) & "%"
                    If (Double.TryParse(reader("POL_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("POL_SECCO") = dr("POL_SECCO") + "(*)"
                        End If
                    End If

                    dr("COV_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COV_SECCO"))
                    dr("COV_AVAIL") = String.Format("{0:0.00}", reader("COV_AVAIL")) & "%"
                    If (Double.TryParse(reader("COV_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("COV_SECCO") = dr("COV_SECCO") + "(*)"
                        End If
                    End If

                    dr("FUMI_SECCO") = String.Format(nfi, "{0:n2}", reader("FUMI_SECCO"))
                    dr("FUMI_AVAIL") = String.Format("{0:0.00}", reader("FUMI_AVAIL")) & "%"
                    If (Double.TryParse(reader("FUMI_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("FUMI_SECCO") = dr("FUMI_SECCO") + "(*)"
                        End If
                    End If

                    If (reader("TIPO_DATO").ToString().Contains("VLE")) Then
                        dr("FUMI_SECCO") = "-"
                    End If

                    If (reader("TIPO_DATO").ToString().Contains("SUPERI")) Then
                        dr("SO2_SECCO") = String.Format("{0:n0}", reader("SO2_SECCO"))
                    End If

                    dr("TIPO_DATO") = reader("TIPO_DATO").ToString()
                    dt.Rows.Add(dr)
                    dr = dt.NewRow()

                Catch ex As Exception
                    Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                    Continue While
                End Try

            End While

        End If


        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        connection.Close()


        Return dt

    End Function

    Private Function GetSecondBollaTable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 3                                                  'In this case the getData is splitted in two part so the first 3 steps was executed by the first part
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim aia As Int32 = 1
        Dim dataType As String = " AND TIPO_DATO LIKE '%MAX_ORE%' ORDER BY INS_ORDER"
        Dim methodName As String = GetCurrentMethod()

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

        dt.Columns.Add(New Data.DataColumn("SO2_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("CO_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("NOX_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("POL_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POL_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("COV_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("COV_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("FUMI_SECCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("FUMI_AVAIL", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("TIPO_DATO", GetType(String)))


        Dim logStatement As String = "SELECT * FROM [ARPA_WEB_CONCENTRAZIONI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & dataType
        Dim max_ore As Integer
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)
        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try


        If (reader.HasRows) Then
            While reader.Read()
                max_ore = reader("SO2_SECCO")
            End While
        End If
        reader.Close()

        dataType = " AND TIPO_DATO LIKE '%COUNT%' ORDER BY INS_ORDER"
        logStatement = "SELECT * FROM [ARPA_WEB_CONCENTRAZIONI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & dataType
        Dim count1, count2, count3, count4, count5, count6 As Double
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)
        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try


        If (reader.HasRows) Then
            While reader.Read()
                Try
                    count1 = reader("SO2_SECCO")
                    count2 = reader("CO_SECCO")
                    count3 = reader("NOX_SECCO")
                    count4 = reader("POL_SECCO")
                    count5 = reader("COV_SECCO")
                    count6 = reader("FUMI_SECCO")
                Catch ex As Exception
                    Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                    Continue While
                End Try

            End While

        End If
        reader.Close()

        dataType = " AND TIPO_DATO LIKE '%AVG%' ORDER BY INS_ORDER"
        logStatement = "SELECT * FROM [ARPA_WEB_CONCENTRAZIONI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & dataType
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)
        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try


        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat
        nfi.NumberGroupSeparator = ""
        Dim dr As Data.DataRow = dt.NewRow()
        ' Modifica per stored procedure, se la data selezionata è maggiore del 2018 avvio la divisione per 7 perchè uso l'altra stored procedure
        Dim data As Date = #1/1/2018#
        Dim result As Integer = Date.Compare(startTime, data)
        ' Fine modifica
        If (reader.HasRows) Then
            While reader.Read()
                Try
                    dr("IDX_REPORT") = reader("IDX_REPORT")
                    'dr("INS_ORDER") = String.Format("{0:n0}", reader("INS_ORDER"))
                    If (type = 2) Then
                        dr("ORA") = "Giorno"
                    ElseIf (type = 1) Then
                        dr("ORA") = "Mese"
                    Else
                        dr("ORA") = "Anno"
                    End If


                    If result < 0 Then

                        dr("SO2_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2_SECCO"))
                        dr("SO2_AVAIL") = String.Format("{0:##}", count1 / max_ore * 100) & "%"
                        dr("CO_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO_SECCO"))
                        dr("CO_AVAIL") = String.Format("{0:##}", count2 / max_ore * 100) & "%"
                        dr("NOX_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX_SECCO"))
                        dr("NOX_AVAIL") = String.Format("{0:##}", count3 / max_ore * 100) & "%"
                        dr("POL_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POL_SECCO"))
                        dr("POL_AVAIL") = String.Format("{0:##}", count4 / max_ore * 100) & "%"
                        dr("COV_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COV_SECCO"))
                        dr("COV_AVAIL") = String.Format("{0:##}", count5 / max_ore * 100) & "%"
                        dr("FUMI_SECCO") = String.Format(nfi, "{0:0}", reader("FUMI_SECCO"))
                        dr("FUMI_AVAIL") = String.Format("{0:##}", count6 / max_ore * 100) & "%"

                        If reader("TIPO_DATO").ToString().Contains("AVG") Then          ''Il valore di media non va nello specchietto subito sotto alla tabella principale, non in fondo alla tabella                    
                            dr("TIPO_DATO") = ""
                        Else
                            dr("TIPO_DATO") = reader("TIPO_DATO").ToString()
                        End If

                    Else
                        dr("SO2_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2_SECCO"))
                        dr("SO2_AVAIL") = String.Format("{0:##}", (count1 / max_ore * 100) / 7) & "%"
                        dr("CO_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO_SECCO"))
                        dr("CO_AVAIL") = String.Format("{0:##}", (count2 / max_ore * 100) / 7) & "%"
                        dr("NOX_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX_SECCO"))
                        dr("NOX_AVAIL") = String.Format("{0:##}", (count3 / max_ore * 100) / 7) & "%"
                        dr("POL_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POL_SECCO"))
                        dr("POL_AVAIL") = String.Format("{0:##}", (count4 / max_ore * 100) / 7) & "%"
                        dr("COV_SECCO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COV_SECCO"))
                        dr("COV_AVAIL") = String.Format("{0:##}", (count5 / max_ore * 100) / 7) & "%"
                        dr("FUMI_SECCO") = String.Format(nfi, "{0:0}", reader("FUMI_SECCO"))
                        dr("FUMI_AVAIL") = String.Format("{0:##}", (count6 / max_ore * 100) / 7) & "%"
                        If reader("TIPO_DATO").ToString().Contains("AVG") Then          ''Il valore di media non va nello specchietto subito sotto alla tabella principale, non in fondo alla tabella                    
                            dr("TIPO_DATO") = ""
                        Else
                            dr("TIPO_DATO") = reader("TIPO_DATO").ToString()
                        End If

                    End If

                    dt.Rows.Add(dr)
                    dr = dt.NewRow()

                Catch ex As Exception
                    Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                    Continue While
                End Try

            End While



            dr("ORA") = "Ore Valide"
            dr("SO2_SECCO") = String.Format("{0:n0}", count1)
            dr("CO_SECCO") = String.Format("{0:n0}", count2)
            dr("NOX_SECCO") = String.Format("{0:n0}", count3)
            dr("POL_SECCO") = String.Format("{0:n0}", count4)
            dr("COV_SECCO") = String.Format("{0:n0}", count5)
            dr("FUMI_SECCO") = String.Format("{0:n0}", count6)
            dt.Rows.Add(dr)

        End If

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        connection.Close()

        Return dt

    End Function

    Private Function GetFirstCTETable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, ByVal type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim commandCTE As System.Data.SqlClient.SqlCommand
        Dim connectionCTE As New SqlConnection(connectionStringCTE)
        Dim queryNumber As Integer = 0
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim methodName As String = GetCurrentMethod()
        Dim retLong As Long

        Try
            ' Tenta di aprire la connessione
            connectionCTE.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
            Return dt
        End Try


        If Form1.reportType = 0 Then                                                      ' It was needed thanks to the genius who wrote the logics in the portal :))
            type = 3
        ElseIf Form1.reportType = 1 Then
            type = 2
        End If


        dt.Columns.Add(New Data.DataColumn("INTESTAZIONE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("COT", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("H2O", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2RIF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("MWE", GetType(String)))

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Dim reader As System.Data.SqlClient.SqlDataReader

        Dim storedProcedureName As String = If(startTime >= "01/01/2018", "sp_AQMSNT_FILL_ARPA_REPORT_WEB", "sp_AQMSNT_FILL_ARPA_REPORT_WEB2017")
        Dim testCMD As New Data.SqlClient.SqlCommand(storedProcedureName, connectionCTE)

        If startTime >= "01/01/2018" Then
            testCMD.Parameters.Add("@aia", Data.SqlDbType.Int, 11)
            testCMD.Parameters("@aia").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@aia").Value = Form1.aia
        End If


        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = section

        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = startTime



        testCMD.Parameters.Add("@tipoestrazione", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@tipoestrazione").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@tipoestrazione").Value = type

        testCMD.Parameters.Add("@o2_rif", Data.SqlDbType.Float, 11)
        testCMD.Parameters("@o2_rif").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@o2_rif").Value = O2RefDict(cteConfiguration)

        testCMD.Parameters.Add("@retval", Data.SqlDbType.BigInt, 8)
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

        retLong = testCMD.Parameters("@retval").Value
        hnf = testCMD.Parameters("@HNF").Value.ToString()
        htran = testCMD.Parameters("@H_TRANS").Value.ToString()

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Dim log_statement As String = "SELECT * FROM [ARPA_REPORT_WEB] WHERE IDX_REPORT = " & retLong.ToString() & " ORDER BY N_RIGA"
        commandCTE = New System.Data.SqlClient.SqlCommand(log_statement, connectionCTE)

        Try
            reader = commandCTE.ExecuteReader()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try

        Dim dr As Data.DataRow = dt.NewRow()
        While reader.Read()
            Try
                dr("INTESTAZIONE") = reader("INTESTAZIONE")
                dr("CO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO"))
                dr("NOX") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX"))
                dr("SO2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2"))
                dr("POLVERI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POLVERI"))
                dr("COT") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COT"))
                dr("QFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("QFUMI"))
                dr("O2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("O2"))
                dr("TFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("TFUMI"))
                dr("PFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PFUMI"))
                dr("H2O") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("H2O"))
                dr("O2RIF") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("O2RIF"))
                dr("MWE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("MWE"))
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

    Private Function GetSecondCTETable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, ByVal type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim commandCTE As System.Data.SqlClient.SqlCommand
        Dim connectionCTE As New SqlConnection(connectionStringCTE)
        Dim queryNumber As Integer = 3
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim methodName As String = GetCurrentMethod()
        Dim retLong As Long
        Dim dataType As String = " ORDER BY INS_ORDER"
        Dim avgType As String

        Try
            ' Tenta di aprire la connessione
            connectionCTE.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
            Return dt
        End Try


        If Form1.reportType = 0 Then                                                      ' It was needed thanks to the genius who wrote the logics in the portal :))
            type = 3
            avgType = "Media Annuale"
        ElseIf Form1.reportType = 1 Then
            type = 2
            avgType = "Media Mensile"
        Else
            dt = Nothing
            Return dt
        End If


        dt.Columns.Add(New Data.DataColumn("IDX_REPORT", GetType(Double)))
        dt.Columns.Add(New Data.DataColumn("INS_ORDER", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("ORA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_NOX", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("CO_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_CO", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("SO2_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_SO2", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("POLVERI_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POLVERI_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_POLVERI", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("COT_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("COT_VLE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_COT", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2_MIS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2_RIF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("ORE_NF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("UFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("MWE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUELGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_COT", GetType(String)))

        Dim reader As System.Data.SqlClient.SqlDataReader

        Dim storedProcedureName As String = If(startTime >= "01/01/2018", "sp_AQMSNT_FILL_ARPA_MESE_ANNO_REPORT", "sp_AQMSNT_FILL_ARPA_MESE_ANNO_REPORT2017")
        Dim testCMD As New Data.SqlClient.SqlCommand(storedProcedureName, connectionCTE)

        If startTime >= "01/01/2018" Then
            testCMD.Parameters.Add("@aia", Data.SqlDbType.Int, 11)
            testCMD.Parameters("@aia").Direction = Data.ParameterDirection.Input
            testCMD.Parameters("@aia").Value = Form1.aia
        End If


        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = section

        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = startTime



        testCMD.Parameters.Add("@IS_MESE", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@IS_MESE").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@IS_MESE").Value = Form1.reportType

        testCMD.Parameters.Add("@o2_rif", Data.SqlDbType.Float, 11)
        testCMD.Parameters("@o2_rif").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@o2_rif").Value = O2RefDict(cteConfiguration)

        testCMD.Parameters.Add("@retval", Data.SqlDbType.BigInt, 8)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_NOX", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_NOX").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_CO", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_CO").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_SO2", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_SO2").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_POLVERI", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_POLVERI").Direction = Data.ParameterDirection.Output

        Try
            testCMD.ExecuteScalar()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
            dt = Nothing
            Return dt
        End Try

        vleCo = testCMD.Parameters("@LL_GG_CO").Value.ToString()
        vleNox = testCMD.Parameters("@LL_GG_NOX").Value.ToString()
        retLong = testCMD.Parameters("@retval").Value

        Dim log_statement As String = "SELECT * FROM [ARPA_WEB_MESE_ANNO_REPORT] WHERE IDX_REPORT = " & retLong.ToString() & dataType
        commandCTE = New System.Data.SqlClient.SqlCommand(log_statement, connectionCTE)
        Try
            reader = commandCTE.ExecuteReader()
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
                dr("ORA") = reader("ORA") 'CountString.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX"))
                If (Not ((IsDBNull(reader("O2_RIF"))) And reader("ORA") <> "Media Mensile")) Then
                    Dim current_o2rif As Double
                    If (Not (IsDBNull(reader("O2_RIF")))) Then
                        current_o2rif = Convert.ToDouble(reader("O2_RIF"))
                    Else
                        current_o2rif = 0.0
                    End If

                    If (reader("ORA") = "Media Mensile" Or current_o2rif <> O2RefDict(cteInvertedConfiguration)) Then
                        dr("NOX_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX_IC"))
                        dr("NOX_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("NOX_VLE"))
                        dr("DISP_NOX") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_NOX"))
                        dr("IS_BOLD_NOX") = reader("IS_BOLD_NOX")
                        dr("CO_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO_IC"))
                        dr("CO_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("CO_VLE"))
                        dr("DISP_CO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_CO"))
                        dr("IS_BOLD_CO") = reader("IS_BOLD_CO")
                        dr("SO2_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2_IC"))
                        dr("SO2_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("SO2_VLE"))
                        dr("DISP_SO2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_SO2"))
                        dr("IS_BOLD_SO2") = reader("IS_BOLD_SO2")
                        dr("POLVERI_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POLVERI_IC"))
                        dr("POLVERI_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("POLVERI_VLE"))
                        dr("DISP_POLVERI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_POLVERI"))
                        dr("IS_BOLD_POLVERI") = reader("IS_BOLD_POLVERI")
                        dr("O2_MIS") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("O2_MIS"))
                        dr("O2_RIF") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", O2RefDict(cteConfiguration))
                        dr("TFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("TFUMI"))
                        dr("PFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PFUMI"))
                        dr("ORE_NF") = reader("ORE_NF")
                        dr("QFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("QFUMI"))
                        dr("UFUMI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("UFUMI"))
                        dr("MWE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("MWE"))
                        dr("QGAS") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("QGAS"))
                        dr("QFUELGAS") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("QFUELGAS"))
                        dr("PORTATA_CO") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_CO"))
                        dr("PORTATA_NOX") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_NOX"))
                        dr("PORTATA_SO2") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_SO2"))
                        dr("PORTATA_POLVERI") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_POLVERI"))
                        dr("COT_IC") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COT_IC"))
                        dr("COT_VLE") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("COT_VLE"))
                        dr("DISP_COT") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("DISP_COT"))
                        dr("PORTATA_COT") = String.Format(CultureInfo.CreateSpecificCulture("it-IT"), "{0:n2}", reader("PORTATA_COT"))

                        If (reader("ORA") = "Media Annuale") Then
                            dr("SO2_VLE") = "N.A."
                            dr("POLVERI_VLE") = "N.A."
                            dr("NOX_VLE") = "N.A."
                        End If
                    End If
                End If

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

    Private Sub downloadReportFlussi(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)


        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim templatePath As String
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""
        Dim d2 As Date


        Select Case Form1.reportType
            Case 0
                reportTitle = "152_MASSICO_ANNO_" & startDate.Year.ToString()
                datanh3 = "01/01/2020"
                d2 = New Date(2020, 1, 1)
            Case 1
                d2 = New Date(2020, mesenh3, 1)
            Case 2
                d2 = New Date(2020, mesenh3, 1)
        End Select

        'System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        If (startDate >= d2) Then
            templatePath = Path.Combine(rootPath, "template", "E9_152_MESE_MASS_CAMINI.xls")
        Else
            templatePath = Path.Combine(rootPath, "template", "152_MESE_MASS_CAMINI.xls")

        End If
        wBook = excel.Workbooks.Open(templatePath)
        wSheet = wBook.ActiveSheet()

        Dim i As Integer
        Dim cc As Integer
        Dim kk As Integer
        Dim app As String
        Dim tabspace As Integer
        cc = 11



        Select Case Form1.reportType
            Case 0
                wSheet.Range("NomeTabella").Value = "152 MASSICO ANNUALE CAMINI DI RAFFINERIA"
                wSheet.Range("IntervalloDate").Value = "Report Annuale dell'anno " + startDate.Year.ToString()
                wSheet.Range("B8").Value = "Mese"
                If (startDate >= d2) Then
                    wSheet.Range("NOTA_FRASE").Value = "Parametro NH3 disponibile sul camino E9 dal mese di Ottobre 2020 a seguito del completamento dei test funzionali, in ottemperanza alla prescrizione [43] dell’AIA DM92/2018"
                Else
                    wSheet.Range("NOTA_FRASE").Value = ""
                End If
            Case 1
                wSheet.Range("NomeTabella").Value = "152 MASSICO MENSILE CAMINI DI RAFFINERIA"
                Dim startDateFormatted As DateTime = DateTime.Parse(startDate).Date
                wSheet.Range("IntervalloDate").Value = "Report Mensile del Mese di " & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM yyyy}", startDateFormatted)
                reportTitle = "152_MASSICO_MESE_" & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM_yyyy}", Date.Parse(startDate))
                wSheet.Range("B8").Value = "Giorno"
                wSheet.Range("NOTA_FRASE").Value = ""
        End Select

        wSheet.Range("NomeTabella").Font.Bold = True
        wSheet.Range("NomeCentrale").Value = "ENI R&M Taranto " ' " & MySharedMethod.GetChimneyName(Convert.ToInt16(Sezione.Text.ToString()))
        wSheet.Range("NomeCentrale").Font.Bold = True
        wSheet.Range("SisMisura").Value = "Sistema di Monitoraggio delle Emissioni"
        wSheet.Range("SisMisura").Font.Bold = True
        wSheet.Range("TitoloTabella").Value = reportTitle
        wSheet.Range("TitoloTabella").Font.Bold = True
        wSheet.Range("IntervalloDate").Font.Bold = True

        Dim quit As String
        If (startDate >= d2) Then
            quit = 43
        Else
            quit = 42
        End If

        'riga grigia
        For i = 0 To Form1.dgv.Rows.Count - 2  'Escluse righe VLE, superi, Max e Min
            wSheet.Rows(cc + i).Insert()
        Next
        Dim stringa As String
        If (startDate >= d2) Then
            stringa = "AQ"
        Else
            stringa = "AP"
        End If
        'prima tabella (solo parte esterna e bordi, non inserisce i dati)
        wSheet.Cells(i + 10, 2) = "Giorno"
        ComboStatus.Report(State.TableLoading)
        For i = 0 To Form1.dgv.Rows.Count - 1
            Dim currentRow As DataGridViewRow = Form1.dgv.Rows(i)
            app = "B" & i + 11 & ":" & stringa & i + 11
            wSheet.Range(app).NumberFormat = "@"
            wSheet.Range(app).BorderAround()

            For kk = 2 To quit
                If currentRow.Cells(kk).Value.ToString() = "" Then
                    wSheet.Cells(i + 11, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        wSheet.Cells(i + 11, kk) = String.Format("{0:HH.mm}", currentRow.Cells(kk).Value.ToString())
                    Else
                        If (startDate < d2 And kk >= 38) Then
                            wSheet.Cells(i + 11, kk) = currentRow.Cells(kk + 1).Value.ToString()
                        Else
                            wSheet.Cells(i + 11, kk) = currentRow.Cells(kk).Value.ToString()
                        End If

                    End If
                End If
            Next


        Next



        'codice per  seconda tabella
        tabspace = 11 + Form1.dgv.Rows.Count + 4 ' spazio tra le due tabelle
        'per modificare le righe (quantità, prima erano 6) della seconda tabella(oer valide, ore n.f) modificare il GridView2.Rows.Count e gv_dailyrep.Rows.Count.(prima a +3 e -4)

        For i = Form1.dgv.Rows.Count To Form1.dgv2.Rows.Count + Form1.dgv.Rows.Count + 2
            wSheet.Rows(cc + i + 1).Insert()
        Next
        'la tabella in basso
        For i = 0 To Form1.dgv2.Rows.Count - 3
            Dim currentRow As DataGridViewRow = Form1.dgv2.Rows(i)
            app = "B" & i + tabspace & ":" & stringa & i + tabspace
            wSheet.Range(app).NumberFormat = "@"
            wSheet.Range(app).BorderAround()
            app = "C" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "D" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "E" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "F" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "G" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "H" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "I" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "J" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "K" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "L" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "M" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "N" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "O" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "P" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "Q" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "R" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "S" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "T" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "U" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "V" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "W" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "X" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "Y" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "Z" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AA" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AB" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AC" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AD" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AE" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AF" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            app = "AG" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AH" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AI" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AJ" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AK" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AL" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AM" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AN" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AO" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AP" & i + tabspace
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            If (startDate >= d2) Then
                app = "AQ" & i + tabspace
                wSheet.Range(app).BorderAround()
                wSheet.Range(app).HorizontalAlignment = -4108
                wSheet.Range(app).VerticalAlignment = -4108
            End If


            'lunghezza tabella secondaria
            For kk = 2 To quit
                If currentRow.Cells(kk).Value.ToString() = "" Then
                    wSheet.Cells(i + tabspace, kk) = ""
                Else
                    If (startDate < d2 And kk >= 38) Then
                        wSheet.Cells(i + tabspace, kk) = currentRow.Cells(kk + 1).Value.ToString()
                    Else
                        wSheet.Cells(i + tabspace, kk) = currentRow.Cells(kk).Value.ToString()
                    End If

                End If
            Next
        Next

        Dim tabspacenota As Integer
        ' codice aggiuntivo (tabella flussi di massa, specchietto riassuntivo TUTTI I CAMINI)
        tabspace = 11 + Form1.dgv.Rows.Count + 4
        'per scritta NH3 
        tabspacenota = 34 + Form1.dgv.Rows.Count + 4
        Dim last As String
        Dim last1 As Integer
        'specchietto visibile solo se l'aia 2018 e il report è annuale
        ComboStatus.Report(State.SheetLoading)
        If ((wSheet.Range("B8").Value = "Mese") And (Form1.aia = 1)) Then

            'intestazionne inquinanti ton
            'AG8 AL9
            If (startDate >= d2) Then
                '  wSheet.Range("B8:H9").Copy()
                wSheet.Range("AG8:AL9").Copy()
                last = "I"
                last1 = 9

            Else
                wSheet.Range("B8:G9").Copy()
                last = "G"
                last1 = 7
            End If

            If (startDate >= d2) Then
                wSheet.Range("C34").Select()
            Else
                wSheet.Range("B34").Select()
            End If
            wSheet.Paste()
            If (startDate >= d2) Then
                wSheet.Range("B8:B9").Copy()
                wSheet.Range("B34").Select()
                wSheet.Paste()
                wSheet.Range("H35").Value = "(¹) NH3(Ton)"
                wSheet.Range("C" & tabspacenota).Value = "(¹) NH3: contributo del solo camino E9 (rif. prescrizione n. [43] del PIC Decreto AIA n.92/2018 ) "
                wSheet.Range("C" & tabspacenota + 1).Value = " (²)NOX: contributi camini E1 + E2 + E4 + E7 +E8 + E9 (rif. prescrizione n. [28] del PIC Decreto AIA n.92/2018 ) "
                wSheet.Range("I35").Value = "(²) NOX(Ton) (RIF. BAT 57)"
            End If

            wSheet.Range("C34").Value = "E1+E2+E4+E7+E8+E9+E10"
            If (startDate >= d2) Then
                wSheet.Range("C35").Value = "NOX(Ton)"
                wSheet.Range("D35").Value = "SO2(Ton) (RIF. BAT 58)"
            Else
                wSheet.Range("C35").Value = "NOX(Ton)"
                wSheet.Range("D35").Value = "SO2(Ton)"
            End If


            wSheet.Range("E35").Value = "Polveri(Ton)"
            wSheet.Range("F35").Value = "CO(Ton)"
            wSheet.Range("G35").Value = "COV(Ton)"

            ' wSheet.Range("C34:G35").Borders.Weight = "3"
            '  For p ì 0 To gv?dailyrep.Rows.Count ' 1
            'margini specchietto riassuntivo (mesi e somma annuale)
            For p = 0 To Form1.dgv.Rows.Count - 1
                Dim currentRow As DataGridViewRow = Form1.dgv.Rows(p)
                '          app = "B" & p + 36 & ":" & last & p + 36
                app = "B" & p + 36 & ":" & last & p + 36
                wSheet.Range(app).NumberFormat = "@"
                wSheet.Range(app).BorderAround()
                wSheet.Range(app).Borders.Weight = "2"
                app = "B" & p + 36



                wSheet.Range(app).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                '  wSheet.Range(app).Borders.Weight = "1"



                For kk = 2 To last1


                    If currentRow.Cells(kk).Value.ToString() = "&nbsp;" Then
                        wSheet.Cells(p + 36, kk) = ""
                    Else
                        'riga relativa al vle
                        wSheet.Cells(p + 36, kk).Copy()
                        wSheet.Cells(p + 37, kk).Select()
                        wSheet.Paste()
                        'valori vle
                        If (startDate >= d2) Then
                            wSheet.Cells(p + 37, 3).Value = "N.A."
                        Else
                            wSheet.Cells(p + 37, 3).Value = 700
                        End If


                        wSheet.Cells(p + 37, 4).Value = 2000
                        wSheet.Cells(p + 37, 5).Value = 50
                        wSheet.Cells(p + 37, 6).Value = "N.A"
                        wSheet.Cells(p + 37, 7).Value = "N.A"
                        If (startDate >= d2) Then
                            wSheet.Cells(p + 37, 8).Value = "N.A"
                            wSheet.Cells(p + 37, 9).Value = 700
                        End If
                        '  wSheet.Cells(p + 37, kk).font.bold = True
                        '  Else
                        '
                        If kk = 2 Then ' ORA

                            wSheet.Cells(p + 36, kk) = String.Format("{0:HH.mm}", currentRow.Cells(kk).Value.ToString())
                            wSheet.Cells(p + 36, kk).font.bold = True
                            wSheet.Cells(p + 36, kk).Copy()
                            wSheet.Cells(p + 37, kk).Select()
                            wSheet.Paste()
                            wSheet.Cells(p + 37, kk).Value = "VLE"
                            '  wSheet.Range(p + 37, kk).Borders.Weight = "3"

                        Else
                            'gv_dailyrep.SelectedRow.Cells(kk + 40).Text
                            wSheet.Cells(p + 36, kk) = String.Format("{0:0.00}", (currentRow.Cells(kk + 41).Value.ToString()))

                            'riga grigia somma 
                            If (p = Form1.dgv.Rows.Count - 1) Then
                                wSheet.Cells(p + 36, 2).Interior.color = Color.LightGray
                                wSheet.Cells(p + 36, kk).Interior.color = Color.LightGray
                                wSheet.Cells(p + 37, kk).font.bold = True
                            End If


                        End If
                    End If

                Next

            Next

            'specchietto Camino E3 
            'mesi, somma , vle
            For ep = 0 To Form1.dgv.Rows.Count - 1
                Dim currentRow As DataGridViewRow = Form1.dgv.Rows(ep)
                app = "L" & ep + 36 & ":Q" & ep + 36
                wSheet.Range(app).NumberFormat = "@"
                wSheet.Range(app).BorderAround()



                wSheet.Range(app).Borders.Weight = "2"

                wSheet.Range("B8:G9").Copy()
                wSheet.Range("L34").Select()
                wSheet.Paste()

                wSheet.Range("C35:G35").Copy()
                wSheet.Range("M35").Select()
                wSheet.Paste()
                wSheet.Range("M34").Value = "CAMINO E3"
                wSheet.Range("N35").Value = "SO2(Ton)"
                wSheet.Range("Q35").Value = "COT(Ton)"




                For kk = 12 To 17

                    If currentRow.Cells(kk).Value.ToString() = "" Then
                        wSheet.Cells(ep + 36, kk) = ""
                    Else

                        'riga specifica vle annuali
                        wSheet.Cells(ep + 36, kk).Copy()
                        wSheet.Cells(ep + 37, kk).Select()
                        wSheet.Paste()
                        wSheet.Cells(ep + 37, 12).Value = "VLE"
                        wSheet.Cells(ep + 37, 13).Value = 750
                        wSheet.Cells(ep + 37, 14).Value = 400
                        wSheet.Cells(ep + 37, 15).Value = 10
                        wSheet.Cells(ep + 37, 16).Value = "N.A"
                        wSheet.Cells(ep + 37, 17).Value = "N.A"



                        'mesi in grassetto
                        If kk = 12 Then ' ORA
                            wSheet.Cells(ep + 36, kk) = String.Format("{0:HH.mm}", currentRow.Cells(2).Value.ToString())
                            wSheet.Cells(ep + 36, kk).font.bold = True
                            wSheet.Cells(ep + 36, kk).Copy()
                            wSheet.Cells(ep + 37, kk).Select()
                            wSheet.Paste()
                        Else
                            'il nuemro viene mostrato solo con due cifre decimali
                            Dim doubleValue As Double
                            If (Double.TryParse(currentRow.Cells(kk).Value.ToString(), doubleValue)) Then
                                wSheet.Cells(ep + 36, kk) = String.Format("{0:0.00}", (doubleValue) / 1000)
                            Else
                                wSheet.Cells(ep + 36, kk) = currentRow.Cells(kk).Value.ToString()
                            End If



                            'colore grigio riga somma annuale

                            If (ep = Form1.dgv.Rows.Count - 1) Then
                                wSheet.Cells(ep + 36, 12).Interior.color = Color.LightGray
                                wSheet.Cells(ep + 36, kk).Interior.color = Color.LightGray

                                wSheet.Cells(ep + 37, kk).font.bold = True
                            End If

                        End If
                        'End If
                    End If
                Next
            Next
        End If

        Dim cellOffset As Integer = tabspace
        ' Dim stringa As String

        For i = 3 To Form1.dgv2.Rows.Count - 1

            Dim currentRow As DataGridViewRow = Form1.dgv2.Rows(i)
            'righe nella tabella i-1 elimina una riga
            If (Form1.reportType = 1 And ((currentRow.Cells(2).Value.ToString().Contains("Sup")) Or (currentRow.Cells(2).Value.ToString().Contains("VLE")))) Then
                Continue For
            End If

            Dim dontMerge As Boolean
            dontMerge = (currentRow.Cells(2).Value.ToString().Contains("Totale"))

            app = "B" & i + cellOffset - 1
            wSheet.Range(app).BorderAround()
            app = "B" & i + cellOffset & ":" & stringa & i + cellOffset

            For Each cell In wSheet.Range(app).Cells
                wSheet.Range(cell, cell).BorderAround()
            Next

            app = "C" & i + cellOffset & ":G" & i + cellOffset
            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "H" & i + cellOffset & ":L" & i + cellOffset
            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "M" & i + cellOffset & ":Q" & i + cellOffset
            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "R" & i + cellOffset & ":V" & i + cellOffset
            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "W" & i + cellOffset & ":AA" & i + cellOffset
            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            app = "AB" & i + cellOffset & ":AF" & i + cellOffset
            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            If (startDate >= d2) Then
                app = "AG" & i + cellOffset & ":AL" & i + cellOffset
            Else
                app = "AG" & i + cellOffset & ":AK" & i + cellOffset
            End If

            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108
            If (startDate >= d2) Then
                app = "AM" & i + cellOffset & ":AQ" & i + cellOffset
            Else
                app = "AL" & i + cellOffset & ":AP" & i + cellOffset
            End If
            wSheet.Range(app).BorderAround()
            If (Not (dontMerge)) Then
                wSheet.Range(app).Merge()
            End If
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108


            'tabella secondaria 43
            'riga)

            For kk = 2 To 43
                If currentRow.Cells(kk).Value.ToString() = "" Then
                    wSheet.Cells(i + cellOffset, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        wSheet.Cells(i + cellOffset, kk) = String.Format("{0:HH.mm}", currentRow.Cells(kk).Value.ToString())
                    Else
                        If (startDate < d2 And kk = 38) Then
                            wSheet.Cells(i + cellOffset, kk) = currentRow.Cells(kk + 1).Value.ToString()
                        Else
                            wSheet.Cells(i + cellOffset, kk) = currentRow.Cells(kk).Value.ToString()
                        End If

                    End If
                End If
            Next
        Next



        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        excel.DisplayAlerts = False
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        ComboStatus.Report(State.FinishedReport)
        wBook.Close()
        excel.DisplayAlerts = True
        excel.Quit()

        Marshal.ReleaseComObject(wSheet)
        Marshal.ReleaseComObject(wBook)

        Marshal.ReleaseComObject(excel)
        wSheet = Nothing
        wBook = Nothing
        excel = Nothing
        MySharedMethod.KillAllExcels()

        If (startDate = endDate) Then

            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If



    End Sub

    Private Sub downloadReportBolla(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""
        Dim d2 As Date = New Date(2020, 1, 1)

        If (startDate >= d2) Then

            wBook = excel.Workbooks.Open(Path.Combine(rootPath, "template", "BAT_152_GIORNO_BOLLA_CAMINI.xls"))
        Else

            wBook = excel.Workbooks.Open(Path.Combine(rootPath, "template", "152_GIORNO_BOLLA_CAMINI.xls"))

        End If

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        wSheet = wBook.ActiveSheet()

        Select Case Form1.reportType
            Case 0
                wSheet.Range("NomeTabella").Value = "152 CONCENTRAZIONI ANNUALE CAMINI DI RAFFINERIA"
                wSheet.Range("IntervalloDate").Value = "Report Annuale dell'anno " + String.Format(New CultureInfo("it-IT", False), "{0:yyyy}", DateTime.Parse(startDate, New CultureInfo("it-IT", False)))
                wSheet.Cells(2, 8).Value = "MESE"
                reportTitle = "152_BOLLA_ANNO_" & startDate.Year.ToString()
            Case 1

                wSheet.Range("NomeTabella").Value = "152 CONCENTRAZIONI MENSILI CAMINI DI RAFFINERIA"
                Dim startDateFormatted As DateTime = DateTime.Parse(startDate).Date
                wSheet.Range("IntervalloDate").Value = "Report Mensile del Mese di " & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM yyyy}", startDateFormatted)
                wSheet.Cells(2, 8).Value = "GIORNO"
                reportTitle = "152_BOLLA_MESE_" & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM_yyyy}", Date.Parse(startDate))
        End Select



        wSheet.Range("NomeTabella").Font.Bold = True
        wSheet.Range("NomeCentrale").Value = "ENI R&M Taranto " ' " & MySharedMethod.GetChimneyName(Convert.ToInt16(Sezione.Text.ToString()))
        wSheet.Range("NomeCentrale").Font.Bold = True
        wSheet.Range("SisMisura").Value = "Sistema di Monitoraggio delle Emissioni"
        wSheet.Range("SisMisura").Font.Bold = True
        wSheet.Range("TitoloTabella").Value = reportTitle
        wSheet.Range("TitoloTabella").Font.Bold = True

        Dim i As Integer
        Dim z As Integer
        Dim insert_tab As Integer
        Dim cc As Integer
        Dim kk As Integer
        Dim app As String
        Dim col As Integer
        Dim tabcounter As Integer
        Dim colgv As Integer
        cc = 11

        ' Inserisci le righe nel foglio di lavoro Excel
        ' Inserimento righe per la prima tabella
        For i = 0 To Form1.dgv.Rows.Count - 1 - 5 ' Escludi righe VLE, superi, Max e Min
            wSheet.Rows(cc + i).Insert()
        Next

        ' Imposta "Giorno" nella cella specificata
        wSheet.Cells(i + 10, 2) = "Giorno"

        ComboStatus.Report(State.TableLoading)
        ' Popolazione della prima tabella
        For i = 0 To Form1.dgv.Rows.Count - 1 - 5
            Form1.dgv.ClearSelection()
            Form1.dgv.Rows(i).Selected = True ' Seleziona la riga corrente

            app = "B" & i + 11 & ":N" & i + 11
            wSheet.Range(app).NumberFormat = "@"
            wSheet.Range(app).BorderAround()

            app = "C" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "D" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "E" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "F" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "G" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "H" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "I" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "J" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "K" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "L" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "M" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "N" & i + 11
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            For kk = 2 To 14
                If Form1.dgv.Rows(i).Cells(kk).Value.ToString() = "" Then
                    wSheet.Cells(i + 11, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        wSheet.Cells(i + 11, kk) = String.Format("{0:HH.mm}", Form1.dgv.Rows(i).Cells(kk).Value)
                    Else
                        wSheet.Cells(i + 11, kk) = Form1.dgv.Rows(i).Cells(kk).Value.ToString()
                    End If
                End If
            Next
        Next

        ComboStatus.Report(State.SheetLoading)
        Dim cellOffset As Integer = 11 + 2

        ' Seconda parte del ciclo per la popolazione della prima tabella
        For i = Math.Max(Form1.dgv.Rows.Count, 0) To Form1.dgv.Rows.Count - 1
            Form1.dgv.ClearSelection()
            Form1.dgv.Rows(i).Selected = True ' Seleziona la riga corrente
            If (((Form1.dgv.Rows(i).Cells(2).Value.ToString().Contains("Sup")) Or (Form1.dgv.Rows(i).Cells(2).Value.ToString().Contains("VLE")))) Then
                Continue For
            End If

            app = "B" & i + cellOffset
            wSheet.Range(app).BorderAround()

            app = "C" & i + cellOffset & ":D" & i + cellOffset
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).Merge()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "E" & i + cellOffset & ":F" & i + cellOffset
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).Merge()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "G" & i + cellOffset & ":H" & i + cellOffset
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).Merge()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "I" & i + cellOffset & ":J" & i + cellOffset
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).Merge()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "K" & i + cellOffset & ":L" & i + cellOffset
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).Merge()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            app = "M" & i + cellOffset & ":N" & i + cellOffset
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).Merge()
            wSheet.Range(app).HorizontalAlignment = -4108
            wSheet.Range(app).VerticalAlignment = -4108

            For kk = 2 To 14
                If Form1.dgv.Rows(i).Cells(kk).Value.ToString() = "" Then
                    wSheet.Cells(i + cellOffset, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        wSheet.Cells(i + cellOffset, kk) = String.Format("{0:HH.mm}", Form1.dgv.Rows(i).Cells(kk).Value)
                    Else
                        wSheet.Cells(i + cellOffset, kk) = Form1.dgv.Rows(i).Cells(kk).Value.ToString()
                    End If
                End If
            Next
        Next

        ' Popolazione della seconda tabella
        col = 2
        insert_tab = i + cc + 4

        For z = 0 To Form1.dgv2.Rows.Count - 1
            Form1.dgv2.ClearSelection()
            Form1.dgv2.Rows(z).Selected = True ' Seleziona la riga corrente

            colgv = 1
            For tabcounter = 2 To Form1.dgv2.Columns.Count
                If Form1.dgv2.Rows(z).Cells(tabcounter - 1).Value.ToString() = "" Then 'Or Form1.dgv2.Rows(z).Cells(tabcounter - 1).Value.ToString().Contains("AVG")
                    wSheet.Cells(insert_tab, colgv) = ""
                Else
                    wSheet.Cells(insert_tab, colgv) = Form1.dgv2.Rows(z).Cells(tabcounter - 1).Value.ToString()
                End If
                colgv += 1
            Next

            insert_tab += 1
            col += 1
        Next

        Form1.dgv2.ClearSelection() ' Deseleziona eventuali righe selezionate


        excel.DisplayAlerts = False
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        excel.DisplayAlerts = False
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.PageSetup.PrintArea = "A1:N" & insert_tab
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        wBook.Close()
        excel.DisplayAlerts = True
        excel.Quit()

        Marshal.ReleaseComObject(wSheet)
        Marshal.ReleaseComObject(wBook)

        Marshal.ReleaseComObject(excel)
        wSheet = Nothing
        wBook = Nothing
        excel = Nothing
        MySharedMethod.KillAllExcels()
        ComboStatus.Report(State.FinishedReport)

        If (startDate = endDate) Then
            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If

    End Sub

    Private Sub downloadYearlyReportCamini(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""
        Dim d2 As Date = New Date(2020, mesenh3, 1)

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        If (startDate.Year >= d2.Year And Form1.section = 6) Then

            wBook = excel.Workbooks.Open(Path.Combine(rootPath, "template", "E9_152_CONC_ANNO_TARANTO_RAFF_COV.xls"))
        Else

            wBook = excel.Workbooks.Open(Path.Combine(rootPath, "template", "152_CONC_ANNO_TARANTO_RAFF_COV.xls"))

        End If

        wSheet = wBook.ActiveSheet()

        Dim percentuale As String

        Dim i As Integer
        Dim j As Integer
        Dim cc As Integer
        Dim app As String
        Dim col As Integer
        Dim insert_tab As Integer
        cc = 11

        ComboStatus.Report(State.TableLoading)

        wSheet.Range("NomeTabella").Value = "152_CONC_ANNO"
        wSheet.Range("NomeTabella").Font.Bold = True
        wSheet.Range("NomeCentrale").Value = "ENI R&M - Raffineria di Taranto - CAMINO " & MySharedMethod.GetChimneyName(Convert.ToInt16(Form1.section.ToString()))
        wSheet.Range("NomeCentrale").Font.Bold = True
        wSheet.Range("SisMisura").Value = "Sistema di Monitoraggio delle Emissioni"
        wSheet.Range("SisMisura").Font.Bold = True

        wSheet.Range("TitoloTabella").Font.Bold = True
        wSheet.Range("IntervalloDate").Value = "Report Annuale Anno " & startDate.Year
        wSheet.Range("IntervalloDate").Font.Bold = True
        wSheet.Range("HNF").Value = hnf
        wSheet.Range("HNF").Font.Bold = True
        wSheet.Range("HTRANS").Value = htran
        wSheet.Range("HTRANS").Font.Bold = True
        reportTitle = MySharedMethod.GetChimneyName(Convert.ToInt16(Form1.section.ToString())) & "_CONC_ANNO_" & startDate.Year

        If (startDate.Year < 2018) Then
            percentuale = "- Dlgs 152 (70%)"
        ElseIf (startDate.Year = 2018) Then
            percentuale = ""
            wSheet.Range("Gestione70").Value = "Dal 1/01/2018  al 31/10/2018 le medie orarie sono validate con disponibilità 70%."
            wSheet.Range("Gestione75").Value = " Dal 1/11/2018 le medie orarie sono validate con disponibilità al 75%."
        Else
            percentuale = "- Dlgs 152 (75%)"
        End If

        wSheet.Range("TitoloTabella").Value = "Report Annuale concentrazioni medie mensili (Nox,Co, SO2, Polveri, COV)  " & percentuale.ToString()
        If (startDate.Year >= d2.Year And Form1.section = 6) Then
            wSheet.Range("TitoloTabella").Value = "Report Annuale concentrazioni medie mensili (Nox,Co, SO2, Polveri, COV,NH3)  " & percentuale.ToString()
            wSheet.Range("NOTA_E9").Value = "Parametro NH3 disponibile sul camino E9 dal mese di Ottobre 2020 a seguito del completamento dei test funzionali, in ottemperanza alla prescrizione [43] dell’AIA DM92/2018"
        End If

        If (Form1.section <> 2) Then
            Try
                wSheet.Range("NOTA_E2").Value = ""
            Catch ex As Exception
            End Try
        End If

        If (Form1.aia = 1) Then
            If (startDate.Year > "2018") Then
                Try
                    wSheet.Range("NOTA_FRASE2").Value = ""
                    wSheet.Range("NOTA_FRASE").Value = ""
                Catch ex As Exception
                End Try
            ElseIf (startDate.Year = "2018") Then
                Try
                    wSheet.Range("NOTA_FRASE2").Value = ""
                Catch ex As Exception
                End Try
            End If
        ElseIf (Form1.aia = 0) Then
            If (startDate.Year = "2018") Then
                Try
                    wSheet.Range("NOTA_FRASE").Value = ""
                Catch ex As Exception
                End Try
            End If
        End If

        Dim firstRow As Integer
        Dim firstColumn As String
        Dim lastColumn As String
        Dim currentExcelCol As String

        firstRow = wSheet.Range("FirstRow").Row
        firstColumn = wSheet.Range("FIRST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
        lastColumn = wSheet.Range("LAST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)

        ComboStatus.Report(State.SheetLoading)

        For i = 0 To Form1.dgv2.Rows.Count - 1
            ' Seleziona la riga corrente in DataGridView
            Form1.dgv2.ClearSelection()
            Form1.dgv2.Rows(i).Selected = True

            app = firstColumn & (i + firstRow).ToString & ":" & lastColumn & (i + firstRow).ToString
            wSheet.Rows(firstRow + i + 1).Insert()
            wSheet.Range(firstColumn & firstRow & ":" & lastColumn & firstRow).Copy(wSheet.Range(app))

            If ((i <> Form1.dgv2.Rows.Count - 1) And (i <> 0)) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
            End If
            If (i = Form1.dgv2.Rows.Count - 1) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End If

            For j = 0 To Form1.dgv2.Columns.Count - 1
                ' Rimpiazza BoundField con DataGridViewColumn e accedi a HeaderText o DataPropertyName
                Dim dataField As String = Form1.dgv2.Columns(j).DataPropertyName
                If dataField.StartsWith("IS_BOLD") Then
                    Dim inquinante = dataField.Split({"IS_BOLD_"}, StringSplitOptions.RemoveEmptyEntries)(0)
                    currentExcelCol = wSheet.Range(inquinante + "_IC").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
                    app = currentExcelCol + Convert.ToString(i + firstRow) + ":" + currentExcelCol + Convert.ToString(i + firstRow)

                    If Not ((Form1.section = 2) And (inquinante = "SO2")) Then

                        If Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) = 1 Then
                            wSheet.Range(app).Font.Bold = True
                            wSheet.Range(app).Interior.Color = Color.Red
                            wSheet.Range(app).Font.Color = Color.White
                        Else
                            wSheet.Range(app).Interior.Color = Color.White
                            wSheet.Range(app).Font.Bold = False
                            wSheet.Range(app).Font.Color = Color.Black
                        End If

                    End If
                End If

                ' Se non c'è il nome della colonna nel template corrispondente al nome sul DataGrid, salta la scrittura
                Try
                    col = wSheet.Range(dataField).Column
                Catch ex As Exception
                    Continue For
                End Try

                If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "" Then
                    wSheet.Cells(i + firstRow, col) = ""
                Else
                    wSheet.Cells(i + firstRow, col) = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                End If
            Next
        Next

        ' Specchietto in basso (report mensile)
        insert_tab = wSheet.Range("FIRSTROW_SUMMARY").Row

        For z = 0 To Form1.dgv.Rows.Count - 1
            Form1.dgv.ClearSelection()
            Form1.dgv.Rows(z).Selected = True

            For j = 0 To Form1.dgv.Columns.Count - 1
                Try
                    col = wSheet.Range("SUMM_" + Form1.dgv.Columns(j).DataPropertyName).Column
                Catch ex As Exception
                    Continue For
                End Try

                If Form1.dgv.Rows(z).Cells(j).Value.ToString() = "" Then
                    wSheet.Cells(insert_tab + z, col) = ""
                Else
                    wSheet.Cells(insert_tab + z, col) = Form1.dgv.Rows(z).Cells(j).Value.ToString()
                End If
            Next
        Next
        Form1.dgv.ClearSelection()

        ComboStatus.Report(State.FinishedReport)
        excel.DisplayAlerts = False
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        excel.DisplayAlerts = False
        wSheet.PageSetup.LeftMargin = Double.Parse(ConfigurationManager.AppSettings("LeftMargin").ToString)
        wSheet.PageSetup.RightMargin = Double.Parse(ConfigurationManager.AppSettings("RightMargin").ToString)
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        wBook.Close()
        excel.DisplayAlerts = True
        excel.Quit()

        Marshal.ReleaseComObject(wSheet)
        Marshal.ReleaseComObject(wBook)

        Marshal.ReleaseComObject(excel)
        wSheet = Nothing
        wBook = Nothing
        excel = Nothing
        MySharedMethod.KillAllExcels()
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")
        ComboStatus.Report(State.FinishedReport)
        If (startDate = endDate) Then

            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If



    End Sub

    Private Sub downloadMonthlyReportCamini(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""
        Dim d2 As Date = New Date(2020, mesenh3, 1)
        Dim templateName As String = ""
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")



        If Form1.section = 6 And Form1.aia = 1 And Form1.startDate >= d2 Then
            templateName = "E9_152_CONC_MESE_TARANTO_RAFF_COV.xls"
        ElseIf Form1.section = 7 Then
            templateName = "152_CONC_MESE_TARANTO_RAFF_COV_NO_GIC.xls"
        ElseIf Form1.section = 3 OrElse Form1.section = 4 Then
            If Form1.aia = 0 Then
                templateName = "152_CONC_MESE_TARANTO_RAFF_COV_NO_GIC.xls"
            ElseIf Form1.aia = 1 Then
                templateName = "4_7152_CONC_MESE_TARANTO_RAFF_COV.xls"
            End If
        Else
            templateName = "152_CONC_MESE_TARANTO_RAFF_COV.xls"
        End If

        wBook = excel.Workbooks.Open(Path.Combine(rootPath, "template", templateName))
        wSheet = wBook.ActiveSheet()

        Dim percentuale As String

        Dim DataCambioPercentuale As Date
        Dim DataScelta As Date

        DataCambioPercentuale = New Date(2018, 11, 1)

        DataScelta = New Date(startDate.Year, startDate.Month, 1)
        Dim compara As Integer = DateTime.Compare(DataCambioPercentuale, DataScelta)



        If (compara = 0 Or compara < 0) Then
            percentuale = " "

        Else
            percentuale = "- Dlgs 152 (70%)"
        End If

        wSheet.Range("NomeTabella").Value = "152_CONC_MESE"
        wSheet.Range("NomeTabella").Font.Bold = True
        wSheet.Range("NomeCentrale").Value = "ENI R&M - Raffineria di Taranto - CAMINO " & MySharedMethod.GetChimneyName(Form1.section.ToString())
        wSheet.Range("NomeCentrale").Font.Bold = True
        wSheet.Range("SisMisura").Value = "Sistema di Monitoraggio delle Emissioni"
        wSheet.Range("SisMisura").Font.Bold = True
        wSheet.Range("TitoloTabella").Value = "Report Mensile concentrazioni medie giornaliere (Nox,Co,So2,Polveri,Cov) " & percentuale.ToString()
        If (startDate > datanh3 And Form1.section = 6) Then
            wSheet.Range("TitoloTabella").Value = "Report Mensile concentrazioni medie giornaliere (Nox,Co,So2,Polveri,Cov ,NH3) " & percentuale.ToString()

        End If

        wSheet.Range("TitoloTabella").Font.Bold = True
        Dim startDateFormatted As DateTime = DateTime.Parse(startDate).Date
        wSheet.Range("IntervalloDate").Value = "Report Mensile del Mese di " & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM yyyy}", startDateFormatted)
        wSheet.Range("IntervalloDate").Font.Bold = True

        wSheet.Range("HNF").Value = hnf
        wSheet.Range("HNF").Font.Bold = True
        wSheet.Range("HTRANS").Value = htran
        wSheet.Range("HTRANS").Font.Bold = True
        reportTitle = MySharedMethod.GetChimneyName(Convert.ToInt16(Form1.section.ToString())) & "_CONC_" & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM_yyyy}", Date.Parse(startDate))
        If (Form1.section <> 2) Then
            Try
                wSheet.Range("NOTA_E2").Value = ""
            Catch ex As Exception

            End Try
        End If

        Dim i As Integer
        Dim j As Integer

        Dim app As String
        Dim col As Integer
        Dim insert_tab As Integer
        Dim firstRow As Integer
        Dim firstColumn As String
        Dim lastColumn As String
        Dim currentExcelCol As String

        firstRow = wSheet.Range("FirstRow").Row
        firstColumn = wSheet.Range("FIRST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
        lastColumn = wSheet.Range("LAST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)

        ComboStatus.Report(State.TableLoading)

        For i = 0 To Form1.dgv2.Rows.Count - 1
            ' Seleziona la riga corrente
            Form1.dgv2.ClearSelection()
            Form1.dgv2.Rows(i).Selected = True

            app = firstColumn & (i + firstRow).ToString & ":" & lastColumn & (i + firstRow).ToString
            wSheet.Rows(firstRow + i + 1).Insert()
            wSheet.Range(firstColumn & firstRow & ":" & lastColumn & firstRow).Copy(wSheet.Range(app))

            If ((i <> Form1.dgv2.Rows.Count - 1) And (i <> 0)) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
            End If
            If (i = Form1.dgv2.Rows.Count - 1) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End If

            For j = 0 To Form1.dgv2.Columns.Count - 1
                ' Rileva il nome della colonna tramite il tag o il nome del DataGridViewColumn
                Dim columnName As String = Form1.dgv2.Columns(j).Name

                If columnName.StartsWith("IS_BOLD") Then
                    Dim inquinante = columnName.Split({"IS_BOLD_"}, StringSplitOptions.RemoveEmptyEntries)(0)
                    currentExcelCol = wSheet.Range(inquinante + "_IC").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
                    app = currentExcelCol + Convert.ToString(i + firstRow) + ":" + currentExcelCol + Convert.ToString(i + firstRow)

                    If Not ((Form1.section = 2) And (inquinante = "SO2")) Then

                        If Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) = 1 Then
                            wSheet.Range(app).Font.Bold = True
                            wSheet.Range(app).Interior.Color = Color.Red
                            wSheet.Range(app).Font.Color = Color.White
                        Else
                            wSheet.Range(app).Interior.Color = Color.White
                            wSheet.Range(app).Font.Bold = False
                            wSheet.Range(app).Font.Color = Color.Black
                        End If

                    End If
                End If

                ' SE NON C'E' IL NOME DI COLONNA SUL TEMPLATE CORRISPONDENTE AL NOME SUL DATAGRID SALTA LA SCRITTURA SU TEMPLATE
                Try
                    col = wSheet.Range(columnName).Column
                Catch ex As Exception
                    Continue For
                End Try

                Dim cellText As String = If(Form1.dgv2.Rows(i).Cells(j).Value Is Nothing, "", Form1.dgv2.Rows(i).Cells(j).Value.ToString())

                If cellText = "" Then
                    wSheet.Cells(i + firstRow, col) = ""
                Else
                    wSheet.Cells(i + firstRow, col) = cellText
                End If
            Next
        Next

        ' specchietto in basso (report mensile)
        insert_tab = wSheet.Range("FIRSTROW_SUMMARY").Row
        ComboStatus.Report(State.SheetLoading)
        For z = 0 To Form1.dgv.Rows.Count - 1
            Form1.dgv.ClearSelection()
            Form1.dgv.Rows(z).Selected = True
            For j = 0 To Form1.dgv.Columns.Count - 1
                ' SE NON C'E' IL NOME DI COLONNA SUL TEMPLATE CORRISPONDENTE AL NOME SUL DATAGRID SALTA LA SCRITTURA SU TEMPLATE
                Try
                    col = wSheet.Range("SUMM_" + Form1.dgv.Columns(j).Name).Column
                Catch ex As Exception
                    Continue For
                End Try

                Dim cellText As String = If(Form1.dgv.Rows(z).Cells(j).Value Is Nothing, "", Form1.dgv.Rows(z).Cells(j).Value.ToString())

                If cellText = "" Then
                    wSheet.Cells(insert_tab + z, col) = ""
                Else
                    wSheet.Cells(insert_tab + z, col) = cellText
                End If
            Next
        Next

        ComboStatus.Report(State.FinishedReport)
        excel.DisplayAlerts = False
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        excel.DisplayAlerts = False
        wSheet.PageSetup.LeftMargin = Double.Parse(ConfigurationManager.AppSettings("LeftMargin").ToString)
        wSheet.PageSetup.RightMargin = Double.Parse(ConfigurationManager.AppSettings("RightMargin").ToString)
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        wBook.Close()
        excel.DisplayAlerts = True
        excel.Quit()

        Marshal.ReleaseComObject(wSheet)
        Marshal.ReleaseComObject(wBook)

        Marshal.ReleaseComObject(excel)
        wSheet = Nothing
        wBook = Nothing
        excel = Nothing
        MySharedMethod.KillAllExcels()
        ComboStatus.Report(State.FinishedReport)

        If (startDate = endDate) Then

            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()

        End If


    End Sub

    Private Sub downloadYearlyReportCTE(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""
        Dim cteConfigurationString As String
        Dim cteInvertedConfigurationString As String

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        If (cteConfiguration = "cogenerativo") Then
            cteConfigurationString = "ASSETTO COGENERATIVO - O2 AL 15%"
            cteInvertedConfigurationString = "Caldaie (O2 al 3%)"
        Else
            cteConfigurationString = "ASSETTO CALDAIE - O2 AL 3%"
            cteInvertedConfigurationString = "Cogenerativo (O2 al 15%)"
        End If

        wBook = excel.Workbooks.Open(Path.Combine(rootPath, "template", "152_CONC_ANNO_TARANTO.xls"))
        wSheet = wBook.ActiveSheet()

        Dim percentuale As String
        Dim i As Integer
        Dim j As Integer
        Dim cc As Integer
        Dim app As String
        Dim col As Integer
        Dim insert_tab As Integer
        cc = 11
        Dim ci As System.Globalization.CultureInfo
        ci = System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")

        ComboStatus.Report(State.TableLoading)

        wSheet.Range("NomeTabella").Value = "152_CONC_ANNO"
        wSheet.Range("NomeTabella").Font.Bold = True
        wSheet.Range("NomeCentrale").Value = "ENI R&M - Raffineria di Taranto - Camino E3" & Chr(10) & cteConfigurationString
        wSheet.Range("NomeCentrale").Font.Bold = True
        wSheet.Range("SisMisura").Value = "Sistema di Monitoraggio delle Emissioni"
        wSheet.Range("SisMisura").Font.Bold = True

        wSheet.Range("TitoloTabella").Font.Bold = True
        wSheet.Range("IntervalloDate").Value = "Report Annuale Anno " & Date.Parse(startDate, ci).Year
        wSheet.Range("IntervalloDate").Font.Bold = True
        reportTitle = "E3_" & "_CONC_ANNO_" & startDate.Year

        Dim year As Integer = Date.Parse(startDate, ci).Year

        If year < 2018 Then
            percentuale = "- Dlgs 152 (70%)"
        ElseIf year = 2018 Then
            percentuale = ""
            wSheet.Range("Gestione70").Value = "Dal 1/01/2018 al 31/10/2018 le medie orarie sono validate con disponibilità 70%."
            wSheet.Range("Gestione75").Value = "Dal 1/11/2018 le medie orarie sono validate con disponibilità al 75%."
        Else
            percentuale = "- Dlgs 152 (75%)"
        End If


        wSheet.Range("TitoloTabella").Value = "Report Annuale concentrazioni medie mensili (NOX ,CO ,SO2 ,POLVERI, COT)" & percentuale.ToString()






        wSheet.Range("HNF").Value = hnf
        wSheet.Range("HNF").Font.Bold = True
        wSheet.Range("HTRANS").Value = htran
        wSheet.Range("HTRANS").Font.Bold = True
        wSheet.Range("C10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("F10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("I10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("L10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("O10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("W8").Value = "Portata Fumi  anidra a " & O2RefDict(cteConfiguration) & "% di O2 (Nm3/h)"

        wSheet.Range("B28").Value = wSheet.Range("B28").Value & cteInvertedConfiguration

        Dim firstRow As Integer
        Dim firstColumn As String
        Dim lastColumn As String
        Dim currentExcelCol As String


        If (Form1.aia = 1) And (Date.Parse(startDate, New System.Globalization.CultureInfo("it-IT")).Year > "2018") Then
            Try
                wSheet.Range("NOTA_FRASE2").Value = ""
                wSheet.Range("NOTA_FRASE").Value = ""
            Catch ex As Exception
            End Try
        End If

        If (Form1.aia = 1) And (Date.Parse(startDate, New System.Globalization.CultureInfo("it-IT")).Year = "2018") Then
            Try
                wSheet.Range("NOTA_FRASE2").Value = ""
            Catch ex As Exception
            End Try
        Else
            Try
                wSheet.Range("NOTA_FRASE").Value = ""
            Catch ex As Exception
            End Try

        End If

        firstRow = wSheet.Range("FirstRow").Row
        firstColumn = wSheet.Range("FIRST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
        lastColumn = wSheet.Range("LAST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)

        ComboStatus.Report(State.SheetLoading)

        For i = 0 To Form1.dgv2.Rows.Count - 1

            ' Seleziona la riga corrente
            Form1.dgv2.Rows(i).Selected = True

            app = firstColumn & (i + firstRow).ToString() & ":" & lastColumn & (i + firstRow).ToString()
            wSheet.Rows(firstRow + i + 1).Insert()
            wSheet.Range(firstColumn & firstRow & ":" & lastColumn & firstRow).Copy(wSheet.Range(app))

            ' Aggiunge bordi sottili tranne per la prima e ultima riga
            If ((i <> Form1.dgv2.Rows.Count - 1) And (i <> 0)) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
            End If

            ' Aggiunge bordo spesso per l'ultima riga
            If (i = Form1.dgv2.Rows.Count - 1) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End If
            ' Itera sulle colonne
            For j = 0 To Form1.dgv2.Columns.Count - 1
                ' Controlla se la colonna inizia con "IS_BOLD"
                If Form1.dgv2.Columns(j).DataPropertyName.StartsWith("IS_BOLD") Then
                    Dim inquinante = Form1.dgv2.Columns(j).DataPropertyName.Split({"IS_BOLD_"}, StringSplitOptions.RemoveEmptyEntries)(0)
                    currentExcelCol = wSheet.Range(inquinante + "_IC").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
                    app = currentExcelCol + Convert.ToString(i + firstRow) + ":" + currentExcelCol + Convert.ToString(i + firstRow)

                    ' Se la cella è vuota ("&nbsp;"), salta
                    If Convert.ToString(Form1.dgv2.Rows(i).Cells(j).Value) = "" Then
                        Exit For
                    Else

                        If Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) = 1 Then
                            wSheet.Range(app).Font.Bold = Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString())
                            wSheet.Range(app).Interior.Color = Color.Red
                            wSheet.Range(app).Font.Color = Color.White
                        Else
                            wSheet.Range(app).Font.Bold = Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString())
                            wSheet.Range(app).Interior.Color = Color.White
                            wSheet.Range(app).Font.Color = Color.Black

                        End If

                    End If
                End If

                ' Prova a ottenere la colonna in Excel, se fallisce continua il ciclo
                Try
                    col = wSheet.Range(Form1.dgv2.Columns(j).DataPropertyName).Column
                Catch ex As Exception
                    Continue For
                End Try

                ' Se la cella è vuota ("&nbsp;"), scrivi una stringa vuota
                If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "" Then
                    wSheet.Cells(i + firstRow, col) = ""
                Else
                    wSheet.Cells(i + firstRow, col) = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                End If
            Next
        Next

        insert_tab = wSheet.Range("FIRSTROW_SUMMARY").Row

        For z = 0 To Form1.dgv.Rows.Count - 1
            Form1.dgv.Rows(z).Selected = True
            For j = 0 To Form1.dgv.Columns.Count - 1
                ' Prova a ottenere la colonna in Excel, se fallisce continua il ciclo
                Try
                    col = wSheet.Range("SUMM_" + Form1.dgv.Columns(j).DataPropertyName).Column
                Catch ex As Exception
                    Continue For
                End Try

                ' Se la cella è vuota ("&nbsp;"), scrivi una stringa vuota
                If Form1.dgv.Rows(z).Cells(j).Value.ToString() = "&nbsp;" Then
                    wSheet.Cells(insert_tab + z, col) = ""
                Else
                    wSheet.Cells(insert_tab + z, col) = Form1.dgv.Rows(z).Cells(j).Value.ToString()
                End If
            Next
        Next

        ComboStatus.Report(State.FinishedReport)
        excel.DisplayAlerts = False
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        excel.DisplayAlerts = False
        wSheet.PageSetup.LeftMargin = Double.Parse(ConfigurationManager.AppSettings("LeftMargin").ToString)
        wSheet.PageSetup.RightMargin = Double.Parse(ConfigurationManager.AppSettings("RightMargin").ToString)
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        ComboStatus.Report(State.FinishedReport)
        wBook.Close()
        excel.DisplayAlerts = True
        excel.Quit()

        Marshal.ReleaseComObject(wSheet)
        Marshal.ReleaseComObject(wBook)

        Marshal.ReleaseComObject(excel)
        wSheet = Nothing
        wBook = Nothing
        excel = Nothing
        MySharedMethod.KillAllExcels()
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")

        If (startDate = endDate) Then

            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If


    End Sub

    Private Sub downloadMonthlyReportCTE(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""
        Dim cteConfigurationString As String
        Dim cteInvertedConfigurationString As String

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        If (cteConfiguration = "cogenerativo") Then
            cteConfigurationString = "ASSETTO COGENERATIVO - O2 AL 15%"
            cteInvertedConfigurationString = "Caldaie (O2 al 3%)"
        Else
            cteConfigurationString = "ASSETTO CALDAIE - O2 AL 3%"
            cteInvertedConfigurationString = "Cogenerativo (O2 al 15%)"
        End If

        wBook = excel.Workbooks.Open(Path.Combine(rootPath, "template", "152_CONC_MESE_TARANTO.xls"))
        wSheet = wBook.ActiveSheet()

        Dim percent As String = " "
        Dim dateToCompare As Date = New Date(2018, 11, 1)

        If (DateTime.Compare(dateToCompare, startDate)) > 0 Then
            percent = "- Dlgs 152 (70%)"
        End If

        ComboStatus.Report(State.TableLoading)
        wSheet.Range("NomeTabella").Value = "152_CONC_MESE"
        wSheet.Range("NomeTabella").Font.Bold = True
        wSheet.Range("NomeCentrale").Value = "ENI R&M - Raffineria di Taranto - Camino E3" & Chr(10) & cteConfigurationString
        wSheet.Range("NomeCentrale").Font.Bold = True
        wSheet.Range("SisMisura").Value = "Sistema di Monitoraggio delle Emissioni"
        wSheet.Range("SisMisura").Font.Bold = True
        wSheet.Range("TitoloTabella").Value = "Report Mensile concentrazioni medie  giornaliere (NOX, CO, SO2, POLVERI, COT) " & percent
        wSheet.Range("TitoloTabella").Font.Bold = True
        Dim startDateFormatted As DateTime = DateTime.Parse(startDate).Date
        wSheet.Range("IntervalloDate").Value = "Report Mensile del Mese di " & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM yyyy}", startDateFormatted)
        wSheet.Range("IntervalloDate").Font.Bold = True
        reportTitle = "E3_" & "CONC_MESE_" & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM_yyyy}", Date.Parse(startDate))
        wSheet.Range("HNF").Value = hnf
        wSheet.Range("HNF").Font.Bold = True
        wSheet.Range("HTRANS").Value = htran
        wSheet.Range("HTRANS").Font.Bold = True
        wSheet.Range("C10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("F10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("I10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("L10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("O10").Value = "NORM IC a " & O2RefDict(cteConfiguration) & "% di O2 QAL2"
        wSheet.Range("W8").Value = "Portata Fumi  anidra a " & O2RefDict(cteConfiguration) & "% di O2 (Nm3/h)"

        wSheet.Range("B28").Value = wSheet.Range("B28").Value & cteInvertedConfiguration

        Dim i As Integer
        Dim j As Integer
        Dim cc As Integer = 11
        Dim app As String
        Dim col As Integer
        Dim insert_tab As Integer
        Dim firstRow As Integer
        Dim firstColumn As String
        Dim lastColumn As String


        firstRow = wSheet.Range("FirstRow").Row
        firstColumn = wSheet.Range("FIRST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
        lastColumn = wSheet.Range("LAST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)

        ComboStatus.Report(State.SheetLoading)
        For i = 0 To Form1.dgv2.Rows.Count - 1

            app = "B" & (i + cc).ToString & ":AF" & (i + cc).ToString
            wSheet.Rows(cc + i + 2).Insert()
            wSheet.Range("B" & cc & ":AF" & cc).Copy(wSheet.Range(app))

            If ((i <> Form1.dgv2.Rows.Count - 1) And (i <> 0)) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
            End If
            If (i = Form1.dgv2.Rows.Count - 1) Then
                wSheet.Range(app).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End If
            col = 2

            For j = 2 To 37 - 1


                If j = 6 Then

                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        app = "C" + Convert.ToString(i + 11) + ":C" + Convert.ToString(i + 11)

                        If Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) = 1 Then

                            wSheet.Range(app).Font.Bold = Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString())
                            wSheet.Range(app).Interior.Color = Color.Red
                            wSheet.Range(app).Font.Color = Color.White

                        End If

                    End If

                    j = j + 1


                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        Dim doubleVal As Double = 0
                        If Double.TryParse(Form1.dgv2.Rows(i).Cells(j).Value.ToString(), doubleVal) Then
                            wSheet.Cells(i + 11, col) = doubleVal
                        Else
                            wSheet.Cells(i + 11, col) = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        End If

                    End If



                    col = col + 1
                ElseIf j = 10 Then
                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        app = "F" + Convert.ToString(i + 11) + ":F" + Convert.ToString(i + 11)

                        If Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) = 1 Then

                            wSheet.Range(app).Font.Bold = Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString())
                            wSheet.Range(app).Interior.Color = Color.Red
                            wSheet.Range(app).Font.Color = Color.White

                        End If

                    End If
                    j = j + 1
                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        Dim doubleVal As Double = 0
                        If Double.TryParse(Form1.dgv2.Rows(i).Cells(j).Value.ToString(), doubleVal) Then
                            wSheet.Cells(i + 11, col) = doubleVal
                        Else
                            wSheet.Cells(i + 11, col) = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        End If
                    End If
                    col = col + 1
                ElseIf j = 14 Then
                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        app = "I" + Convert.ToString(i + 11) + ":I" + Convert.ToString(i + 11)

                        If Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) = 1 Then

                            wSheet.Range(app).Font.Bold = Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString())
                            wSheet.Range(app).Interior.Color = Color.Red
                            wSheet.Range(app).Font.Color = Color.White

                        End If

                    End If
                    j = j + 1
                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        Dim doubleVal As Double = 0
                        If Double.TryParse(Form1.dgv2.Rows(i).Cells(j).Value.ToString(), doubleVal) Then
                            wSheet.Cells(i + 11, col) = doubleVal
                        Else
                            wSheet.Cells(i + 11, col) = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        End If
                    End If
                    col = col + 1
                ElseIf j = 18 Then
                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        app = "L" + Convert.ToString(i + 11) + ":L" + Convert.ToString(i + 11)

                        If Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) = 1 Then

                            wSheet.Range(app).Font.Bold = Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j).Value.ToString())
                            wSheet.Range(app).Interior.Color = Color.Red
                            wSheet.Range(app).Font.Color = Color.White

                        End If

                    End If
                    j = j + 1
                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        Dim doubleVal As Double = 0
                        If Double.TryParse(Form1.dgv2.Rows(i).Cells(j).Value.ToString(), doubleVal) Then
                            wSheet.Cells(i + 11, col) = doubleVal
                        Else
                            wSheet.Cells(i + 11, col) = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        End If
                    End If
                    col = col + 1
                Else
                    If Form1.dgv2.Rows(i).Cells(j).Value Is Nothing OrElse String.IsNullOrEmpty(Form1.dgv2.Rows(i).Cells(j).Value.ToString()) Then
                        wSheet.Cells(i + 11, col) = ""
                    Else
                        Dim doubleVal As Double = 0
                        If Double.TryParse(Form1.dgv2.Rows(i).Cells(j).Value.ToString(), doubleVal) Then
                            wSheet.Cells(i + 11, col) = doubleVal
                        Else
                            wSheet.Cells(i + 11, col) = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        End If
                    End If
                    col = col + 1 'sposta i dati delle colonne
                End If



                If j < 16 Then
                    wSheet.Cells(i + 12, j).BorderAround()

                End If
                'Tabella sintesi
            Next
        Next

        firstRow = wSheet.Range("FirstRow").Row
        firstColumn = wSheet.Range("FIRST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)
        lastColumn = wSheet.Range("LAST_COLUMN").Address.Split({"$"c}, StringSplitOptions.RemoveEmptyEntries)(0)



        insert_tab = wSheet.Range("FIRSTROW_SUMMARY").Row
        '
        For z = 0 To Form1.dgv.Rows.Count - 1
            Form1.dgv.Rows(z).Selected = True
            For j = 0 To Form1.dgv.Columns.Count - 1
                'SE NON C'E' IL NOME DI COLONNA SUL TEMPLATE CORRISPONDENTE AL NOME SUL DATAGRID SALTA LA SCRITTURA SU TEMPLATE
                Try
                    col = wSheet.Range("SUMM_" + Form1.dgv.Columns(j).DataPropertyName).Column
                Catch ex As Exception
                    Continue For
                End Try

                If Form1.dgv.Rows(z).Cells(j).Value.ToString() = "&nbsp;" Then
                    wSheet.Cells(insert_tab + z, col) = ""
                Else
                    wSheet.Cells(insert_tab + z, col) = Form1.dgv.Rows(z).Cells(j).Value.ToString()
                End If

            Next
        Next


        excel.DisplayAlerts = False
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        excel.DisplayAlerts = False
        wSheet.PageSetup.LeftMargin = Double.Parse(ConfigurationManager.AppSettings("LeftMargin").ToString)
        wSheet.PageSetup.RightMargin = Double.Parse(ConfigurationManager.AppSettings("RightMargin").ToString)
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)

        wBook.Close()
        excel.DisplayAlerts = True
        excel.Quit()

        Marshal.ReleaseComObject(wSheet)
        Marshal.ReleaseComObject(wBook)

        Marshal.ReleaseComObject(excel)
        wSheet = Nothing
        wBook = Nothing
        excel = Nothing
        MySharedMethod.KillAllExcels()
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")
        ComboStatus.Report(State.FinishedReport)

        If (startDate = endDate) Then

            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If

    End Sub


    '    Private Sub DisableForm()

    '        For Each ctrl As Control In Controls
    '            If (Not ctrl.Equals(dgv) And (Not ctrl.Name = ProgressBar1.Name Or Not ctrl.Name = TextBox1.Name)) Then
    '                ctrl.Enabled = False
    '            End If
    '        Next

    '    End Sub

    Private Sub ShowCompletionDialog()                                                                                  ' Crea un'istanza del form modale e la mostra in modalità                                       

        Dim completedDownloadForm As New Form2()
        completedDownloadForm.ShowDialog()

    End Sub


    Private Function GetCurrentMethod() As String
        Dim stackTrace As New StackTrace()
        Dim method As MethodBase = stackTrace.GetFrame(1).GetMethod()
        Return method.Name
    End Function

    Private Sub preRenderFirstTable(section As Integer)

        For Each column As DataGridViewColumn In Form1.dgv.Columns
            If hiddenColumns.Contains(column.DataPropertyName) Then
                UpdateDgvColumnVisibility(False, Form1.dgv, column.Name)
            End If
        Next

        If hiddenColumns.Count.Equals(0) Then
            For Each column As DataGridViewColumn In Form1.dgv.Columns
                If section = 8 Then
                    If column.DataPropertyName = "E9Q_NH3" Or column.DataPropertyName = "NH3_SOMMA" Or column.DataPropertyName = "NOX57_SOMMA" Then
                        UpdateDgvColumnVisibility(True, Form1.dgv, column.Name)
                    End If
                Else
                    UpdateDgvColumnVisibility(True, Form1.dgv, column.Name)
                End If

            Next
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

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
