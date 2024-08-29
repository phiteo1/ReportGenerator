Imports System.Threading
Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices

Public Class Form1

    Dim connectionString As String
    Dim culture As System.Globalization.CultureInfo
    Dim reportType As Int32 = 255
    Dim section As Int32 = 255
    Dim ret As Int32
    Dim ret2 As Int32
    Dim dgv As DataGridView
    Dim dgv2 As DataGridView
    Dim datanh3 As String = ConfigurationManager.AppSettings("datanh3")
    Dim mesenh3 As Integer = ConfigurationManager.AppSettings("mesenh3")
    Dim hiddenColumns As New List(Of String)()
    Dim d2 As Date
    Dim bolla As Byte
    Dim aia As Int32 = 1

    Enum State                  'State Machine of the downloading process
        DataLoading = 1
        TableLoading = 2
        SheetLoading = 3
        FinishedReport = 4
        Finished = 5
    End Enum

    Dim actualState As Byte


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load                                'Inizialitation of the database connection, form's item and of the grid view 

        connectionString = ConfigurationManager.ConnectionStrings("GLOBAL_CONN_STR").ConnectionString
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        TextBox1.Visible = False
        DateTimePicker1.Value = Date.Now.AddYears(-1)
        culture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        culture.NumberFormat.NumberGroupSeparator = ""
        SetDataGridView()


    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DisableForm()
        TextBox1.Visible = True
        reportType = ComboBox2.SelectedIndex
        section = GetSection(ComboBox1.SelectedItem)
        Dim startDate As New DateTime(DateTimePicker1.Value.Year, 1, 1)
        Dim endDate As New DateTime(DateTimePicker2.Value.Year, 1, 1)
        ProgressBar1.Visible = True
        ProgressBar1.Maximum = 100

        'Refresh the GUI when a change in the progress bar occours
        Dim barProgress As New Progress(Of Integer)(Sub(v)
                                                        ProgressBar1.Value = v
                                                    End Sub)

        'Refresh the GUI when a change in the state occours
        Dim StatusProgress As New Progress(Of Integer)(Sub(index)
                                                           Select Case index
                                                               Case 1
                                                                   TextBox1.Text = "Data Loading..."
                                                                   actualState = State.DataLoading
                                                               Case 2
                                                                   TextBox1.Text = "Table Creation..."
                                                                   ProgressBar1.Visible = False
                                                                   actualState = State.TableLoading
                                                               Case 3
                                                                   TextBox1.Text = "Sheet Creation..."
                                                                   actualState = State.SheetLoading
                                                               Case 4
                                                                   TextBox1.Text = "Report year " & startDate.Year.ToString & " downloaded succesfully"
                                                                   actualState = State.FinishedReport
                                                               Case 5
                                                                   TextBox1.Text = "Report generation finished!"
                                                                   actualState = State.Finished
                                                                   Button1.Text = "Generate Again"
                                                                   EnableForm()
                                                                   Me.Hide()
                                                           End Select
                                                       End Sub)
        Dim dataTable1 As DataTable
        Dim dataTable2 As DataTable
        Controls.Add(dgv)
        Controls.Add(dgv2)

        If Not CheckBox1.Checked Then
            aia = 0
        End If

        While (startDate <= endDate)
            ProgressBar1.Value = 0
            If (Not ProgressBar1.Visible) Then
                ProgressBar1.Visible = True
            End If
            If section = 8 Then
                If bolla = 0 Then
                    dataTable1 = Await Task.Run(Function() GetDataFlussi(barProgress, startDate, endDate, section, reportType, 1, dgv))   'Get the data from the database and assign to first data table structure. The function is runned in an other trhead in order to allow the GUI to refresh properly
                    dataTable2 = Await Task.Run(Function() GetDataFlussi(barProgress, startDate, endDate, section, reportType, 2, dgv2))  'Get the data from the database and assign to second data table structure
                Else
                    dataTable1 = Await Task.Run(Function() GetDataBolla1(barProgress, startDate, endDate, section, reportType, dgv))   'Get the data from the database and assign to first data table structure. The function is runned in an other trhead in order to allow the GUI to refresh properly
                    dataTable2 = Await Task.Run(Function() GetDataBolla2(barProgress, startDate, endDate, section, reportType, dgv2))  'Get the data from the database and assign to second data table structure
                End If

            Else
                dataTable1 = Nothing
                dataTable2 = Nothing
            End If

            If dataTable1 Is Nothing OrElse dataTable1.Rows.Count = 0 Then
                MessageBox.Show("Nessun dato restituito o errore nella query", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                EnableForm()
                Return
            Else
                dgv.DataSource = dataTable1                                                                                     'Bind the data to the first DataGridView
            End If
            If dataTable2 Is Nothing OrElse dataTable2.Rows.Count = 0 Then
                MessageBox.Show("Nessun dato restituito o errore nella query.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                EnableForm()
                Return
            Else
                dgv2.DataSource = dataTable2                                                                                    'Bind the data to the second DataGridView
            End If

            dgv.Visible = True
            dgv.Visible = False                                                                                             'Dont' worry about that. It's an hack to get the correct number of rows
            dgv2.Visible = True
            dgv2.Visible = False
            TextBox1.Text = "Data Loading..."
            TextBox1.Visible = True

            If bolla = 0 Then
                Await Task.Run(Sub() downloadReportFlussi(StatusProgress, startDate, endDate))                              'Download the reports of the selected years. The function is runned in an other trhead in order to allow the GUI to refresh properly 

            ElseIf bolla = 1 Then
                Await Task.Run(Sub() downloadReportBolla(StatusProgress, startDate, endDate))
            Else
                'todo
            End If

            Dim deltaTime As String
            If (reportType = 0) Then
                deltaTime = "yyyy"                                                                                          'Add one year or one month according to the report type choosed
            Else
                deltaTime = "m"

            End If
            startDate = DateAdd(deltaTime, 1, startDate)
        End While

        Me.Show()
        If Me.WindowState = FormWindowState.Minimized Then
            Me.WindowState = FormWindowState.Normal
        End If
        Me.BringToFront()
        Me.Activate()

    End Sub

    Private Function GetSection(camino As String) As Int32

        Select Case camino

            Case "Camino E1"
                Return 1
            Case "Camino E2"
                Return 2
            Case "Camino E3"
                Return 8
            Case "Camino E4"
                Return 3
            Case "Camino E7"
                Return 4
            Case "Camino E8"
                Return 5
            Case "Camino E9"
                Return 6
            Case "Camino E10"
                Return 7
            Case "Camino E1"
                Return 1
            Case "Flussi di massa"
                bolla = 0
                Return 8
            Case "Bolla di raffineria"
                bolla = 1
                Return 8
            Case Else
                Return 255

        End Select

    End Function

    Private Function GetDataFlussi(progress As IProgress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32, whatTable As Byte, dgv As DataGridView) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim command2 As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim connection2 As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 0
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim dataType As String = " AND TIPO_DATO IS NOT NULL ORDER BY INS_ORDER"


        Try
            ' Tenta di aprire la connessione
            connection.Open()
            connection2.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return dt
        End Try

        Select Case reportType
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
            Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_MASSICI_CAMINI_NODELETE", connection)
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
            testCMD.Parameters("@TIPO_ESTRAZIONE").Value = reportType
            testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
            testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
            Try
                testCMD.ExecuteScalar()
                ret = testCMD.Parameters("@retval").Value
                queryNumber += 3
                progress.Report(queryNumber * progressStep)

                testCMD.Parameters("@idsez").Value = 1
                testCMD.ExecuteScalar()

                ret2 = testCMD.Parameters("@retval").Value
            Catch ex As Exception
                Console.WriteLine("Errore durante l'esecuzione della stored procedure: " & ex.Message, "Errore SQL")
                Return dt
            End Try
            
            dataType = " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"

        End If


        Dim logStatement As String = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & dataType
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)
        Dim reader2 As System.Data.SqlClient.SqlDataReader
        Try
            reader = command.ExecuteReader()
            logStatement = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret2.ToString() & dataType
            command2 = New System.Data.SqlClient.SqlCommand(logStatement, connection2)
            reader2 = command2.ExecuteReader()

        Catch ex As SqlException
            Console.WriteLine("Errore durante l'esecuzione della query: " & ex.Message, "Errore SQL")
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
                    Console.WriteLine("Errore nella lettura dei dati: " & ex.Message)
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


        For Each column As DataGridViewColumn In dgv.Columns
            ' Verifica se il nome della colonna è nella lista delle colonne nascoste
            If hiddenColumns.Contains(column.DataPropertyName) Then
                column.Visible = False
            End If
        Next

        If hiddenColumns.Count = 0 Then
            For Each column As DataGridViewColumn In dgv.Columns
                ' Verifica se il nome della colonna corrisponde ai nomi specificati
                If column.DataPropertyName = "E9Q_NH3" Or column.DataPropertyName = "NH3_SOMMA" Or column.DataPropertyName = "NOX57_SOMMA" Then
                    column.Visible = True
                End If
            Next
        End If

        connection.Close()
        connection2.Close()

        If whatTable = 2 Then
            connection2.Open()
            logStatement = "DELETE FROM ARPA_WEB_MASSICI_CAMINI"

            Using deleteCmd As New SqlCommand(logStatement, connection2)

                Try
                    deleteCmd.ExecuteNonQuery()

                Catch ex As SqlException

                    Console.WriteLine("Errore durante l'esecuzione della query: " & ex.Message, "Errore SQL")

                End Try

            End Using

            connection2.Close()
        End If

        Return dt
    End Function

    Private Function GetDataBolla1(progress As IProgress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32, dgv As DataGridView) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim connection2 As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 0
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount


        Try
            ' Tenta di aprire la connessione
            connection.Open()
            connection2.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            testCMD.Parameters("@aia").Value = aia


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
        testCMD.Parameters("@TIPO_ESTRAZIONE").Value = reportType
        testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
        Try
            testCMD.ExecuteScalar()
        Catch ex As SqlException
            Console.WriteLine("Errore durante l'esecuzione della store procedure: " & ex.Message, "Errore SQL")
        End Try

        ret = testCMD.Parameters("@retval").Value

        queryNumber += 1
        progress.Report(queryNumber * progressStep)


        Dim logStatement As String = "SELECT * FROM [ARPA_WEB_CONCENTRAZIONI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & "  ORDER BY INS_ORDER"
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)

        Try
            reader = command.ExecuteReader()
        Catch ex As Exception
            Console.WriteLine("Errore durante l'esecuzione della query: " & ex.Message, "Errore SQL")
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
                        dr("ORA") = String.Format("{0:MMMM}", reader("ORA"))
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
                    dr("SO2_SECCO") = String.Format("{0:n2}", reader("SO2_SECCO"))
                    dr("SO2_AVAIL") = String.Format("{0:0.00}", reader("SO2_AVAIL")) & "%"
                    If (Double.TryParse(reader("SO2_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("SO2_SECCO") = dr("SO2_SECCO") + "(*)"
                        End If
                    End If


                    dr("CO_SECCO") = String.Format("{0:n2}", reader("CO_SECCO"))
                    dr("CO_AVAIL") = String.Format("{0:0.00}", reader("CO_AVAIL")) & "%"
                    If (Double.TryParse(reader("CO_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("CO_SECCO") = dr("CO_SECCO") + "(*)"
                        End If
                    End If

                    dr("NOX_SECCO") = String.Format("{0:n2}", reader("NOX_SECCO"))
                    dr("NOX_AVAIL") = String.Format("{0:0.00}", reader("NOX_AVAIL")) & "%"
                    If (Double.TryParse(reader("NOX_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("NOX_SECCO") = dr("NOX_SECCO") + "(*)"
                        End If
                    End If

                    dr("POL_SECCO") = String.Format("{0:n2}", reader("POL_SECCO"))
                    dr("POL_AVAIL") = String.Format("{0:0.00}", reader("POL_AVAIL")) & "%"
                    If (Double.TryParse(reader("POL_AVAIL").ToString, availability)) Then
                        If (availability < 70) Then
                            dr("POL_SECCO") = dr("POL_SECCO") + "(*)"
                        End If
                    End If

                    dr("COV_SECCO") = String.Format("{0:n2}", reader("COV_SECCO"))
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
                    Console.WriteLine("Errore nella lettura dei dati: " & ex.Message)
                    Continue While
                End Try

            End While
                
        End If


        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        connection.Close()
        connection2.Close()

        Return dt

    End Function

    Private Function GetDataBolla2(progress As IProgress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32, dgv As DataGridView) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim connection2 As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 3                                                  'In this case the getData is splitted in two part so the first 3 steps was executed by the first part
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim aia As Int32 = 1
        Dim dataType As String = " AND TIPO_DATO LIKE '%MAX_ORE%' ORDER BY INS_ORDER"

        Try
            ' Tenta di aprire la connessione
            connection.Open()
            connection2.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
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


        Dim logStatement As String = "SELECT * FROM [ARPA_WEB_CONCENTRAZIONI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & dataType
        Dim max_ore As Integer
        command = New System.Data.SqlClient.SqlCommand(logStatement, connection)
        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Console.WriteLine("Errore durante l'esecuzione della query: " & ex.Message, "Errore SQL")
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
        Catch ex As Exception
            Console.WriteLine("Errore durante l'esecuzione della query: " & ex.Message, "Errore SQL")
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
                    Console.WriteLine("Errore nella lettura dei dati: " & ex.Message)
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
        Catch ex As Exception
            Console.WriteLine("Errore durante l'esecuzione della query: " & ex.Message, "Errore SQL")
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

                        dr("SO2_SECCO") = String.Format("{0:n2}", reader("SO2_SECCO"))
                        dr("SO2_AVAIL") = String.Format("{0:##}", count1 / max_ore * 100) & "%"
                        dr("CO_SECCO") = String.Format("{0:n2}", reader("CO_SECCO"))
                        dr("CO_AVAIL") = String.Format("{0:##}", count2 / max_ore * 100) & "%"
                        dr("NOX_SECCO") = String.Format("{0:n2}", reader("NOX_SECCO"))
                        dr("NOX_AVAIL") = String.Format("{0:##}", count3 / max_ore * 100) & "%"
                        dr("POL_SECCO") = String.Format("{0:n2}", reader("POL_SECCO"))
                        dr("POL_AVAIL") = String.Format("{0:##}", count4 / max_ore * 100) & "%"
                        dr("COV_SECCO") = String.Format("{0:n2}", reader("COV_SECCO"))
                        dr("COV_AVAIL") = String.Format("{0:##}", count5 / max_ore * 100) & "%"
                        dr("FUMI_SECCO") = String.Format(nfi, "{0:0}", reader("FUMI_SECCO"))
                        dr("FUMI_AVAIL") = String.Format("{0:##}", count6 / max_ore * 100) & "%"

                        If reader("TIPO_DATO").ToString().Contains("AVG") Then          ''Il valore di media non va nello specchietto subito sotto alla tabella principale, non in fondo alla tabella                    
                            dr("TIPO_DATO") = ""
                        Else
                            dr("TIPO_DATO") = reader("TIPO_DATO").ToString()
                        End If

                    Else
                        dr("SO2_SECCO") = String.Format("{0:n2}", reader("SO2_SECCO"))
                        dr("SO2_AVAIL") = String.Format("{0:##}", (count1 / max_ore * 100) / 7) & "%"
                        dr("CO_SECCO") = String.Format("{0:n2}", reader("CO_SECCO"))
                        dr("CO_AVAIL") = String.Format("{0:##}", (count2 / max_ore * 100) / 7) & "%"
                        dr("NOX_SECCO") = String.Format("{0:n2}", reader("NOX_SECCO"))
                        dr("NOX_AVAIL") = String.Format("{0:##}", (count3 / max_ore * 100) / 7) & "%"
                        dr("POL_SECCO") = String.Format("{0:n2}", reader("POL_SECCO"))
                        dr("POL_AVAIL") = String.Format("{0:##}", (count4 / max_ore * 100) / 7) & "%"
                        dr("COV_SECCO") = String.Format("{0:n2}", reader("COV_SECCO"))
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
                    Console.WriteLine("Errore nella lettura dei dati: " & ex.Message)
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

        Return dt

    End Function


    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate > startDate Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate > startDate Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Dim combobox As ComboBox = CType(sender, ComboBox)

        If combobox.SelectedIndex = 1 Then
            DateTimePicker1.CustomFormat = "MMMM yyyy"
            DateTimePicker2.CustomFormat = "MMMM yyyy"

        ElseIf combobox.SelectedIndex = 0 Then
            DateTimePicker1.CustomFormat = "yyyy"
            DateTimePicker2.CustomFormat = "yyyy"

        End If

    End Sub

    Private Sub SetDataGridView()

        dgv = New DataGridView()
        dgv.Visible = False
        dgv.Dock = DockStyle.Fill
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgv.AllowUserToAddRows = False
        dgv.AllowUserToDeleteRows = False
        dgv.AllowUserToResizeRows = False
        dgv.RowHeadersVisible = False
        dgv.Width = 1800
        dgv.AutoGenerateColumns = True



        For Each col As DataGridViewColumn In dgv.Columns
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)
        Next


        dgv2 = New DataGridView()
        dgv2.Visible = False
        dgv2.Dock = DockStyle.Fill
        dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgv2.AllowUserToAddRows = False
        dgv2.AllowUserToDeleteRows = False
        dgv2.AllowUserToResizeRows = False
        dgv2.RowHeadersVisible = False
        dgv2.Width = 1800
        dgv2.AutoGenerateColumns = True



        For Each col As DataGridViewColumn In dgv2.Columns
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.DefaultCellStyle.Font = New Font(dgv2.Font, FontStyle.Bold)
        Next

    End Sub

    Private Sub ShowDataGridView(dataTable As DataTable)
        ' Nascondi tutti i controlli del modulo eccetto la DataGridView
        For Each ctrl As Control In Controls
            If Not ctrl.Equals(dgv) Then
                ctrl.Visible = False
            End If
        Next

        ' Imposta la DataGridView come visibile e imposta i dati
        dgv.DataSource = dataTable


        Controls.Add(dgv)
        dgv.Visible = True
        ' Ridimensiona la DataGridView per occupare tutto lo spazio disponibile
        dgv.Dock = DockStyle.Fill

    End Sub

    Private Sub downloadReportFlussi(ComboStatus As IProgress(Of Integer), startDate As Date, endDate As Date)


        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim templatePath As String
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""



        Select Case reportType
            Case 0
                reportTitle = "152_MASSICO_ANNO_" & startDate.Year.ToString()
                datanh3 = "01/01/2020"
                d2 = New Date(2020, 1, 1)
            Case 1
                d2 = New Date(2020, mesenh3, 1)
            Case 2
                d2 = New Date(2020, mesenh3, 1)
        End Select

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

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



        Select Case reportType
            Case 0
                wSheet.Range("NomeTabella").Value = "152 MASSICO ANNUALE CAMINI DI RAFFINERIA"
                wSheet.Range("IntervalloDate").Value = "Report Annuale dell'anno " + Date.Parse(startDate, New System.Globalization.CultureInfo("it-IT")).Year.ToString()
                wSheet.Range("B8").Value = "Mese"
                If (startDate >= d2) Then
                    wSheet.Range("NOTA_FRASE").Value = "Parametro NH3 disponibile sul camino E9 dal mese di Ottobre 2020 a seguito del completamento dei test funzionali, in ottemperanza alla prescrizione [43] dell’AIA DM92/2018"
                Else
                    wSheet.Range("NOTA_FRASE").Value = ""
                End If
            Case 1
                ' TODO
            Case 2
                ' TODO
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
        Dim stringa As String
        stringa = If(startDate >= d2, "AQ", "AP")
        Dim tabspacenota As Integer
        Dim last As String
        Dim last1 As Integer
        Dim appRange As String

        ' Inserisce righe per la prima tabella
        For i = 0 To dgv.Rows.Count - 2
            wSheet.Rows(cc + i).Insert()
        Next

        ' Prima tabella (parte esterna e bordi, non inserisce dati)
        wSheet.Cells(11, 2) = "Giorno"

        ComboStatus.Report(State.TableLoading)
        For i = 0 To dgv.Rows.Count - 1
            Dim currentRow As DataGridViewRow = dgv.Rows(i)
            Dim rowIndex As Integer = i + 11

            ' Formattazione della riga
            appRange = "B" & rowIndex & ":" & stringa & rowIndex
            With wSheet.Range(appRange)
                .NumberFormat = "@"
                .BorderAround()
            End With

            ' Iterazione colonne per la prima tabella
            For kk = 2 To quit
                Dim cellValue As String = If(currentRow.Cells(kk).Value IsNot Nothing, currentRow.Cells(kk).Value.ToString(), String.Empty)

                If String.IsNullOrEmpty(cellValue) OrElse cellValue = "&nbsp;" Then
                    wSheet.Cells(rowIndex, kk) = ""
                Else
                    If i = 2 Then
                        ' Se è la terza riga, formato ORA
                        Dim cellDateTime As DateTime
                        If DateTime.TryParse(cellValue, cellDateTime) Then
                            wSheet.Cells(rowIndex, kk) = cellDateTime.ToString("HH.mm")
                        Else
                            wSheet.Cells(rowIndex, kk) = cellValue
                        End If
                    Else
                        If startDate < d2 AndAlso kk >= 38 Then
                            wSheet.Cells(rowIndex, kk) = currentRow.Cells(kk + 1).Value.ToString()
                        Else
                            wSheet.Cells(rowIndex, kk) = cellValue
                        End If
                    End If
                End If
            Next
        Next

        tabspace = 11 + dgv.Rows.Count + 4
        tabspacenota = 34 + dgv.Rows.Count + 4



        ' Inserisce righe per la seconda tabella
        For i = dgv.Rows.Count To dgv2.Rows.Count + dgv.Rows.Count + 2
            wSheet.Rows(cc + i + 1).Insert()
        Next

        ' Crea e formatta la seconda tabella
        For i = 0 To dgv2.Rows.Count - 3
            Dim currentRow As DataGridViewRow = dgv2.Rows(i)
            Dim rowIndex As Integer = i + tabspace

            ' Formattazione della riga
            appRange = "B" & rowIndex & ":" & stringa & rowIndex
            With wSheet.Range(appRange)
                .NumberFormat = "@"
                .BorderAround()
            End With

            ' Allineamento celle per colonne specifiche
            Dim columns As String() = {"C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP"}
            If startDate >= d2 Then columns = columns.Concat({"AQ"}).ToArray()

            For Each col In columns
                With wSheet.Range(col & rowIndex.ToString())
                    .BorderAround()
                    .HorizontalAlignment = -4108
                    .VerticalAlignment = -4108
                End With
            Next
            ' Popolamento delle celle nella seconda tabella
            For kk = 2 To quit
                Dim cellValue As String = If(currentRow.Cells(kk).Value IsNot Nothing, currentRow.Cells(kk).Value.ToString(), String.Empty)

                If String.IsNullOrEmpty(cellValue) OrElse cellValue = "&nbsp;" Then
                    wSheet.Cells(rowIndex, kk) = ""
                Else
                    If startDate < d2 AndAlso kk >= 38 Then
                        wSheet.Cells(rowIndex, kk) = currentRow.Cells(kk + 1).Value.ToString()
                    Else
                        wSheet.Cells(rowIndex, kk) = cellValue
                    End If
                End If
            Next
        Next

        ComboStatus.Report(State.TableLoading)
        ' Specchietto riassuntivo, visibile solo se il report è annuale
        If wSheet.Range("B8").Value = "Mese" Then
            ' Intestazione inquinanti tonnellate
            If startDate >= d2 Then
                wSheet.Range("AG8:AL9").Copy()
                last = "I"
                last1 = 9
            Else
                wSheet.Range("B8:G9").Copy()
                last = "G"
                last1 = 7
            End If

            If startDate >= d2 Then
                wSheet.Range("C34").Select()
            Else
                wSheet.Range("B34").Select()
            End If
            wSheet.Paste()

            If startDate >= d2 Then
                wSheet.Range("B8:B9").Copy()
                wSheet.Range("B34").Select()
                wSheet.Paste()
                wSheet.Range("H35").Value = "(¹) NH3(Ton)"
                wSheet.Range("C" & tabspacenota).Value = "(¹) NH3: contributo del solo camino E9 (rif. prescrizione n. [43] del PIC Decreto AIA n.92/2018 )"
                wSheet.Range("C" & tabspacenota + 1).Value = "(²)NOX: contributi camini E1 + E2 + E4 + E7 +E8 + E9 (rif. prescrizione n. [28] del PIC Decreto AIA n.92/2018 )"
                wSheet.Range("I35").Value = "(²) NOX(Ton) (RIF. BAT 57)"
            End If

            wSheet.Range("C34").Value = "E1+E2+E4+E7+E8+E9+E10"
            If startDate >= d2 Then
                wSheet.Range("C35").Value = "NOX(Ton)"
                wSheet.Range("D35").Value = "SO2(Ton) (RIF. BAT 58)"
            Else
                wSheet.Range("C35").Value = "NOX(Ton)"
                wSheet.Range("D35").Value = "SO2(Ton)"
            End If

            wSheet.Range("E35").Value = "Polveri(Ton)"
            wSheet.Range("F35").Value = "CO(Ton)"
            wSheet.Range("G35").Value = "COV(Ton)"

            ' Margini specchietto riassuntivo (mesi e somma annuale)
            For p = 0 To dgv.Rows.Count - 1
                Dim currentRow As DataGridViewRow = dgv.Rows(p)
                appRange = "B" & (p + 36).ToString() & ":" & last & (p + 36).ToString()
                With wSheet.Range(appRange)
                    .NumberFormat = "@"
                    .BorderAround()
                    .Borders.Weight = 2
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                End With

                For kk = 2 To last1
                    Dim cellValue As String = If(currentRow.Cells(kk).Value IsNot Nothing, currentRow.Cells(kk).Value.ToString(), String.Empty)

                    If String.IsNullOrEmpty(cellValue) OrElse cellValue = "&nbsp;" Then
                        wSheet.Cells(p + 36, kk) = ""
                    Else
                        wSheet.Cells(p + 36, kk).Value = cellValue
                        wSheet.Cells(p + 36, kk).Copy()
                        wSheet.Cells(p + 37, kk).PasteSpecial()

                        If startDate >= d2 Then
                            wSheet.Cells(p + 37, 3).Value = "N.A."
                        Else
                            wSheet.Cells(p + 37, 3).Value = 700
                        End If

                        wSheet.Cells(p + 37, 4).Value = 2000
                        wSheet.Cells(p + 37, 5).Value = 50
                        wSheet.Cells(p + 37, 6).Value = "N.A"
                        wSheet.Cells(p + 37, 7).Value = "N.A"

                        If startDate >= d2 Then
                            wSheet.Cells(p + 37, 8).Value = "N.A"
                            wSheet.Cells(p + 37, 9).Value = 700
                        End If

                        If kk = 2 Then ' ORA
                            wSheet.Cells(p + 36, kk).Value = String.Format("{0:HH.mm}", currentRow.Cells(kk).Value)
                            wSheet.Cells(p + 36, kk).Font.Bold = True
                            wSheet.Cells(p + 37, kk).Value = "VLE"
                        Else
                            wSheet.Cells(p + 36, kk).Value = String.Format("{0:0.00}", currentRow.Cells(kk + 41).Value)

                            ' Righe grigie e formato
                            If p = dgv.Rows.Count - 1 Then
                                wSheet.Cells(p + 36, 2).Interior.Color = Color.LightGray
                                wSheet.Cells(p + 36, kk).Interior.Color = Color.LightGray
                                wSheet.Cells(p + 37, kk).Font.Bold = True
                            End If
                        End If
                    End If
                Next
            Next
        End If

        ComboStatus.Report(State.SheetLoading)
        For ep = 0 To dgv.Rows.Count - 1
            ' Ottieni la riga corrente usando l'indice ep
            Dim currentRow As DataGridViewRow = dgv.Rows(ep)

            ' Calcola l'intervallo di celle da formattare
            app = "L" & (ep + 36).ToString() & ":Q" & (ep + 36).ToString()
            wSheet.Range(app).NumberFormat = "@"
            wSheet.Range(app).BorderAround()
            wSheet.Range(app).Borders.Weight = 2 ' Usa numeri per pesi dei bordi, non stringhe

            ' Copia e incolla le aree specifiche
            wSheet.Range("B8:G9").Copy()
            wSheet.Range("L34").PasteSpecial() ' Usa PasteSpecial per maggiore controllo

            wSheet.Range("C35:G35").Copy()
            wSheet.Range("M35").PasteSpecial()

            ' Imposta i valori specifici nelle celle
            wSheet.Range("M34").Value = "CAMINO E3"
            wSheet.Range("N35").Value = "SO2(Ton)"
            wSheet.Range("Q35").Value = "COT(Ton)"

            For kk = 12 To 17
                Dim cellValue As String = String.Empty

                ' Verifica se la cella è nulla o contiene un valore e assegnalo a cellValue
                If currentRow.Cells(kk).Value IsNot Nothing Then
                    cellValue = currentRow.Cells(kk).Value.ToString()
                End If

                If String.IsNullOrEmpty(cellValue) OrElse cellValue = "&nbsp;" Then
                    wSheet.Cells(ep + 36, kk) = ""
                Else
                    ' Copia il valore della cella alla riga successiva
                    wSheet.Cells(ep + 36, kk).Copy()
                    wSheet.Cells(ep + 37, kk).PasteSpecial()

                    ' Imposta valori specifici
                    wSheet.Cells(ep + 37, 12).Value = "VLE"
                    wSheet.Cells(ep + 37, 13).Value = 750
                    wSheet.Cells(ep + 37, 14).Value = 400
                    wSheet.Cells(ep + 37, 15).Value = 10
                    wSheet.Cells(ep + 37, 16).Value = "N.A"
                    wSheet.Cells(ep + 37, 17).Value = "N.A"

                    ' Formattazione per la colonna ORA
                    If kk = 12 Then ' ORA
                        wSheet.Cells(ep + 36, kk).Value = String.Format("{0:HH.mm}", currentRow.Cells(2).Value)
                        wSheet.Cells(ep + 36, kk).Font.Bold = True
                        wSheet.Cells(ep + 36, kk).Copy()
                        wSheet.Cells(ep + 37, kk).PasteSpecial()
                    Else
                        ' Mostra il numero con due cifre decimali
                        wSheet.Cells(ep + 36, kk).Value = String.Format("{0:0.00}", (Convert.ToDouble(currentRow.Cells(kk).Value) / 1000))

                        ' Colore grigio per la riga somma annuale
                        If ep = dgv.Rows.Count - 1 Then
                            wSheet.Cells(ep + 36, 12).Interior.Color = Color.LightGray
                            wSheet.Cells(ep + 36, kk).Interior.Color = Color.LightGray
                            wSheet.Cells(ep + 37, kk).Font.Bold = True
                        End If
                    End If
                End If
            Next
        Next


        Dim cellOffset As Integer = tabspace
        ' Dim stringa As String

        For i = 3 To dgv2.Rows.Count - 1

            Dim currentRow As DataGridViewRow = dgv2.Rows(i)
            'righe nella tabella i-1 elimina una riga
            If (reportType = 1 And ((currentRow.Cells(2).Value.Contains("Sup")) Or (currentRow.Cells(2).Value.Contains("VLE")))) Then
                Continue For
            End If

            Dim dontMerge As Boolean
            dontMerge = (currentRow.Cells(2).Value.Contains("Totale"))

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

                Dim cellValue As String = String.Empty
                If currentRow.Cells(kk).Value IsNot Nothing Then
                    cellValue = currentRow.Cells(kk).Value.ToString()
                Else
                    cellValue = String.Empty
                End If

                If String.IsNullOrEmpty(cellValue) OrElse cellValue = "&nbsp;" Then
                    wSheet.Cells(i + cellOffset, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        wSheet.Cells(i + cellOffset, kk) = String.Format("{0:HH.mm}", currentRow.Cells(kk).Value)
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
        Dim reportPath = Path.Combine(rootPath, "report", reportFileXls)
        Dim reportPathPdf = Path.Combine(rootPath, "report", reportFilePdf)
        excel.DisplayAlerts = False
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.PageSetup.PrintArea = "A1:Z60"
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        ComboStatus.Report(State.FinishedReport)
        If (startDate = endDate) Then
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
            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If



    End Sub

    Private Sub downloadReportBolla(ComboStatus As IProgress(Of Integer), startDate As Date, endDate As Date)

        Dim excel As New Microsoft.Office.Interop.Excel.Application
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

        Select Case reportType
            Case 0
                wSheet.Range("NomeTabella").Value = "152 CONCENTRAZIONI ANNUALE CAMINI DI RAFFINERIA"
                wSheet.Range("IntervalloDate").Value = "Report Annuale dell'anno " + String.Format(New CultureInfo("it-IT", False), "{0:yyyy}", DateTime.Parse(startDate, New CultureInfo("it-IT", False)))
                wSheet.Cells(2, 8).Value = "MESE"
                reportTitle = "152_BOLLA_ANNO_" & startDate.Year.ToString()
            Case 1
                ' TODO
            Case 2
                ' TODO
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
        For i = 0 To dgv.Rows.Count - 1 - 5 ' Escludi righe VLE, superi, Max e Min
            wSheet.Rows(cc + i).Insert()
        Next

        ' Imposta "Giorno" nella cella specificata
        wSheet.Cells(i + 10, 2) = "Giorno"

        ComboStatus.Report(State.TableLoading)
        ' Popolazione della prima tabella
        For i = 0 To dgv.Rows.Count - 1 - 5
            dgv.ClearSelection()
            dgv.Rows(i).Selected = True ' Seleziona la riga corrente

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
                If dgv.Rows(i).Cells(kk).Value.ToString() = "&nbsp;" Then
                    wSheet.Cells(i + 11, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        wSheet.Cells(i + 11, kk) = String.Format("{0:HH.mm}", dgv.Rows(i).Cells(kk).Value)
                    Else
                        wSheet.Cells(i + 11, kk) = dgv.Rows(i).Cells(kk).Value.ToString()
                    End If
                End If
            Next
        Next

        ComboStatus.Report(State.SheetLoading)
        Dim cellOffset As Integer = 11 + 2

        ' Seconda parte del ciclo per la popolazione della prima tabella
        For i = Math.Max(dgv.Rows.Count, 0) To dgv.Rows.Count - 1
            dgv.ClearSelection()
            dgv.Rows(i).Selected = True ' Seleziona la riga corrente

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
                If dgv.Rows(i).Cells(kk).Value.ToString() = "&nbsp;" Then
                    wSheet.Cells(i + cellOffset, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        wSheet.Cells(i + cellOffset, kk) = String.Format("{0:HH.mm}", dgv.Rows(i).Cells(kk).Value)
                    Else
                        wSheet.Cells(i + cellOffset, kk) = dgv.Rows(i).Cells(kk).Value.ToString()
                    End If
                End If
            Next
        Next

        ' Popolazione della seconda tabella
        col = 2
        insert_tab = i + cc + 4

        For z = 0 To dgv2.Rows.Count - 1
            dgv2.ClearSelection()
            dgv2.Rows(z).Selected = True ' Seleziona la riga corrente

            colgv = 1
            For tabcounter = 2 To dgv2.Columns.Count
                If dgv2.Rows(z).Cells(tabcounter - 1).Value.ToString() = "&nbsp;" Then 'Or dgv2.Rows(z).Cells(tabcounter - 1).Value.ToString().Contains("AVG")
                    wSheet.Cells(insert_tab, colgv) = ""
                Else
                    wSheet.Cells(insert_tab, colgv) = dgv2.Rows(z).Cells(tabcounter - 1).Value.ToString()
                End If
                colgv += 1
            Next

            insert_tab += 1
            col += 1
        Next

        dgv2.ClearSelection() ' Deseleziona eventuali righe selezionate


        excel.DisplayAlerts = False
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(rootPath, "report", reportFileXls)
        Dim reportPathPdf = Path.Combine(rootPath, "report", reportFilePdf)
        excel.DisplayAlerts = False
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        wSheet.PageSetup.PrintArea = "A1:Z60"
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportPathPdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        ComboStatus.Report(State.FinishedReport)
        If (startDate = endDate) Then
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
            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If

    End Sub

    Private Sub DisableForm()

        For Each ctrl As Control In Controls
            If (Not ctrl.Equals(dgv) And (Not ctrl.Name = ProgressBar1.Name Or Not ctrl.Name = TextBox1.Name)) Then
                ctrl.Enabled = False
            End If
        Next

    End Sub

    Private Sub EnableForm()

        For Each ctrl As Control In Controls
            If Not ctrl.Equals(dgv) And ctrl.Enabled = False Then
                ctrl.Enabled = True
            End If
        Next

        ResetForm()

    End Sub

    Private Sub ShowCompletionDialog()
        ' Crea un'istanza del form modale
        Dim completedDownloadForm As New Form2()

        ' Mostra il form in modalità modale
        completedDownloadForm.ShowDialog()

    End Sub

    Private Sub ResetForm()

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        Button1.Enabled = True
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        TextBox1.Text = ""
        TextBox1.Visible = False

    End Sub
End Class
