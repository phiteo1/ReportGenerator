Imports System.Threading
Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices

Public Class Form1

    Dim connectionString As String
    Dim culture As System.Globalization.CultureInfo
    Dim reportType As Int32
    Dim section As Int32
    Dim ret As Int32
    Dim ret2 As Int32
    Dim dgv As DataGridView
    Dim datanh3 As String = ConfigurationManager.AppSettings("datanh3")
    Dim mesenh3 As Integer = ConfigurationManager.AppSettings("mesenh3")
    Dim hiddenColumns As New List(Of String)()
    Dim d2 As Date


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        connectionString = ConfigurationManager.ConnectionStrings("GLOBAL_CONN_STR").ConnectionString
        DateTimePicker1.Value = Date.Now.AddYears(-1)
        culture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        culture.NumberFormat.NumberGroupSeparator = ""
        SetDataGridView()


    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        reportType = ComboBox2.SelectedIndex
        section = 8
        Dim startDate As New DateTime(DateTimePicker1.Value.Year, 1, 1)
        Dim endDate As New DateTime(DateTimePicker2.Value.Year, 1, 1)
        ProgressBar1.Location = New Point(465, 501)
        ProgressBar1.Visible = True
        ProgressBar1.Maximum = 100
        Dim progress As New Progress(Of Integer)(Sub(v)
                                                     ProgressBar1.Value = v
                                                 End Sub)

        Dim dataTable As DataTable
        dataTable = Await Task.Run(Function() GetData(progress, startDate, endDate, section, reportType))
        'ShowDataGridView(dataTable)
        dgv.DataSource = dataTable
        'Controls.Add(dgv)
        ProgressBar1.Visible = False
        Button1.Enabled = False
        Button2.Visible = True
        Button3.Visible = True
    End Sub

    Private Function GetData(progress As IProgress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim command2 As System.Data.SqlClient.SqlCommand
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
        dt.Columns.Add(New Data.DataColumn("NOX_SOMMA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2_SOMMA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POLVERI_SOMMA", GetType(String)))


        dt.Columns.Add(New Data.DataColumn("CO_SOMMA", GetType(String)))


        dt.Columns.Add(New Data.DataColumn("COV_SOMMA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NH3_SOMMA", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("NOX57_SOMMA", GetType(String)))

        Select Case reportType
            Case 0
                datanh3 = "01/01/2020"
        End Select

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

        testCMD.ExecuteScalar()
        ret = testCMD.Parameters("@retval").Value
        queryNumber += 3
        progress.Report(queryNumber * progressStep)

        testCMD.Parameters("@idsez").Value = 1
        testCMD.ExecuteScalar()

        ret2 = testCMD.Parameters("@retval").Value
        Dim log_statement As String = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"
        command = New System.Data.SqlClient.SqlCommand(log_statement, connection)

        reader = command.ExecuteReader()
        log_statement = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret2.ToString() & " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"
        command2 = New System.Data.SqlClient.SqlCommand(log_statement, connection2)
        Dim reader2 As System.Data.SqlClient.SqlDataReader
        reader2 = command2.ExecuteReader()
        Dim dr As Data.DataRow = dt.NewRow()
        If (reader.HasRows) Then
            While reader.Read()
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
                dr("NOX_SOMMA") = reader("NOX_SOMMA")
                dr("SO2_SOMMA") = reader("SO2_SOMMA")
                dr("POLVERI_SOMMA") = reader("POLVERI_SOMMA")
                dr("CO_SOMMA") = reader("CO_SOMMA")


                dr("COV_SOMMA") = reader("COV_SOMMA")
                dr("NH3_SOMMA") = reader("NH3_SOMMA")
                dr("NOX57_SOMMA") = reader("NOX57_SOMMA")
                dt.Rows.Add(dr)
                dr = dt.NewRow()


                If (startTime < Date.Parse(datanh3)) Then
                    hiddenColumns.Add("E9Q_NH3")
                    hiddenColumns.Add("NH3_SOMMA")
                    hiddenColumns.Add("NOX57_SOMMA")
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
                        If column.DataPropertyName = "Ed9Q_NH3" Or column.DataPropertyName = "NH3_SOMMA" Or column.DataPropertyName = "NOX57_SOMMA" Then
                            column.Visible = True
                        End If
                    Next
                End If
            End While
            queryNumber += 1
            progress.Report(queryNumber * progressStep)
        End If

        Return dt
    End Function

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate >= startDate Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate >= startDate Then
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Button2.Enabled = False
        Button3.Enabled = False
        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim startDate As New DateTime(DateTimePicker1.Value.Year, 1, 1)
        Dim endDate As New DateTime(DateTimePicker2.Value.Year, 1, 1)
        Dim templatePath As String
        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim reportTitle As String = ""
        


        Select Case reportType
            Case 0
                reportTitle = "152_MASSICO_ANNO"
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
        Controls.Add(dgv)
        dgv.Visible = True
        dgv.Visible = False
        'riga grigia
        For i = 0 To dgv.Rows.Count - 2  'Escluse righe VLE, superi, Max e Min
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
        For i = 0 To dgv.Rows.Count - 1
            ' Non è necessario selezionare esplicitamente la riga, quindi rimuoviamo dgv.SelectedIndex = i

            ' Ottieni la riga corrente
            Dim currentRow As DataGridViewRow = dgv.Rows(i)

            ' Calcola l'intervallo di celle da formattare
            app = "B" & (i + 11).ToString() & ":" & stringa & (i + 11).ToString()
            wSheet.Range(app).NumberFormat = "@"
            wSheet.Range(app).BorderAround()

            ' Itera attraverso le colonne della riga corrente
            For kk = 2 To quit
                Dim cellValue As String
                If currentRow.Cells(kk).Value IsNot Nothing Then
                    cellValue = currentRow.Cells(kk).Value.ToString()
                Else
                    cellValue = String.Empty
                End If
                If String.IsNullOrEmpty(cellValue) OrElse cellValue = "&nbsp;" Then
                    wSheet.Cells(i + 11, kk) = ""
                Else
                    If i = 2 Then ' ORA
                        ' Assumendo che la cella contenga un valore DateTime
                        Dim cellDateTime As DateTime
                        If DateTime.TryParse(cellValue, cellDateTime) Then
                            wSheet.Cells(i + 11, kk) = cellDateTime.ToString("HH.mm")
                        Else
                            wSheet.Cells(i + 11, kk) = cellValue
                        End If
                    Else
                        Dim nextCellValue As String = String.Empty
                        If kk + 1 < currentRow.Cells.Count AndAlso currentRow.Cells(kk + 1).Value IsNot Nothing Then
                            nextCellValue = currentRow.Cells(kk + 1).Value.ToString()
                        End If

                        If startDate < d2 AndAlso kk >= 38 Then
                            wSheet.Cells(i + 11, kk) = nextCellValue
                        Else
                            wSheet.Cells(i + 11, kk) = cellValue
                        End If
                    End If
                End If
            Next
        Next

        Dim tabspacenota As Integer
        ' codice aggiuntivo (tabella flussi di massa, specchietto riassuntivo TUTTI I CAMINI)
        tabspace = 11 + dgv.Rows.Count + 4
        'per scritta NH3 
        tabspacenota = 34 + dgv.Rows.Count + 4
        Dim last As String
        Dim last1 As Integer
        'specchietto visibile solo se l'aia 2018 e il report è annuale
        If ((wSheet.Range("B8").Value = "Mese")) Then

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
            For p = 0 To dgv.Rows.Count - 1
                ' Ottieni la riga corrente usando l'indice p
                Dim currentRow As DataGridViewRow = dgv.Rows(p)

                ' Calcola l'intervallo di celle da formattare
                app = "B" & (p + 36).ToString() & ":" & last & (p + 36).ToString()
                wSheet.Range(app).NumberFormat = "@"
                wSheet.Range(app).BorderAround()
                wSheet.Range(app).Borders.Weight = 2 ' Usa numeri per pesi dei bordi, non stringhe
                app = "B" & (p + 36).ToString()

                wSheet.Range(app).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                For kk = 2 To last1
                    Dim cellValue As String = String.Empty

                    ' Verifica se la cella è nulla o contiene un valore e assegnalo a cellValue
                    If currentRow.Cells(kk).Value IsNot Nothing Then
                        cellValue = currentRow.Cells(kk).Value.ToString()
                    Else
                        cellValue = String.Empty
                    End If

                    If String.IsNullOrEmpty(cellValue) OrElse cellValue = "&nbsp;" Then
                        wSheet.Cells(p + 36, kk) = ""
                    Else
                        ' Copia il valore della cella alla riga successiva
                        wSheet.Cells(p + 36, kk).Copy()
                        wSheet.Cells(p + 37, kk).PasteSpecial()

                        ' Imposta valori specifici
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
                            wSheet.Cells(p + 36, kk).Copy()
                            wSheet.Cells(p + 37, kk).PasteSpecial()
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
        End If

        excel.DisplayAlerts = False
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(rootPath, "report", reportFileXls)
        Dim reportPathPdf = Path.Combine(rootPath, "report", reportFilePdf)
        wBook.SaveAs(reportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
        '  wSheet.PageSetup.PrintArea = "A1:Z60"
        wSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
        wSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        wSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, reportFilePdf, Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False)
        wBook.Close()
        excel.Quit()

        Marshal.ReleaseComObject(wSheet)
        Marshal.ReleaseComObject(wBook)

        Marshal.ReleaseComObject(excel)
        wSheet = Nothing
        wBook = Nothing
        ' wSheet = DBNull.Value
        excel = Nothing
        MySharedMethod.KillAllExcels()
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")
        Button1.Enabled = True
        MsgBox("The Report(s) successfully downloaded. You can find the file(s) in the report directory")

    End Sub
End Class
