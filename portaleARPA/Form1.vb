Imports System.Threading
Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Reflection
Imports System.ComponentModel


Public Class Form1

    Public Shared startDate As Date
    Public Shared endDate As Date
    Public Shared reportType As Int32 = 255
    Public Shared section As Int32 = 255S
    Public Shared dgv As DataGridView
    Public Shared dgv2 As DataGridView
    Public Shared aia As Int32 = 1
    Public Shared isCte As Boolean = False
    Public Shared bolla As Byte
    Dim plant As String
    Dim concretePlant As IImpianto

   
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load                                                                                        'Inizialitation of the database connection, form's item and of the grid view 

        Logger.CreateLogDir()

        plant = ConfigurationManager.AppSettings("Impianto")
        If plant.Contains("ImpiantoTaranto") Then
            concretePlant = New ImpiantoTaranto()
        ElseIf plant.Contains("ImpiantoBrindisi") Then
            concretePlant = New ImpiantoBrindisi()
        End If

        Dim chimneyList As List(Of Camino) = concretePlant.getChimneyList()
        For Each chimney In chimneyList
            ComboBox1.Items.Add(chimney.getName)
        Next


        ComboBox3.Visible = False
        Label6.Visible = False
        ComboBox3.SelectedIndex = 0
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ProgressBar1.Maximum = 100
        TextBox1.Visible = False
        DateTimePicker1.Value = Date.Now.AddYears(-1)
        TextBox1.Text = "Data Loading..."
        SetDataGridView()


    End Sub

    Private Sub Button1_BindingContextChanged(sender As Object, e As EventArgs) Handles Button1.BindingContextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim cteConfiguration As String = ""
        Dim cteInvertedConfiguration As String = ""
        Dim worker As New BackgroundWorker()
        worker.WorkerReportsProgress = False
        worker.WorkerSupportsCancellation = False
        AddHandler worker.DoWork, AddressOf concretePlant.MainThread
        AddHandler worker.RunWorkerCompleted, AddressOf reportCompleted

        reportType = ComboBox2.SelectedIndex
        If (reportType = 0) Then
            startDate = New DateTime(DateTimePicker1.Value.Year, 1, 1)
            endDate = New DateTime(DateTimePicker2.Value.Year, 1, 1)
        ElseIf (reportType = 1) Then
            startDate = New DateTime(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, 1)
            endDate = New DateTime(DateTimePicker2.Value.Year, DateTimePicker2.Value.Month, 1)
        Else
            startDate = New DateTime(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, DateTimePicker1.Value.Day)
            endDate = New DateTime(DateTimePicker2.Value.Year, DateTimePicker2.Value.Month, DateTimePicker2.Value.Day)
        End If

        If Not CheckBox1.Checked Then
            aia = 0
        End If

        If ComboBox3.Visible Then
            isCte = True
        End If

        section = (concretePlant.getChimneyFromName(ComboBox1.SelectedItem)).getSection()
        bolla = (concretePlant.getChimneyFromName(ComboBox1.SelectedItem)).getBolla()
        TextBox1.Visible = True
        Controls.Add(dgv)
        Controls.Add(dgv2)
        dgv.Visible = True                                                                                                                                                 'Dont' worry about that. It's an hack to get the correct number of rows
        dgv.Visible = False
        dgv2.Visible = True
        dgv2.Visible = False

        DisableForm()

        worker.RunWorkerAsync()

    End Sub


    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate.Date >= startDate.Date Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate.Date >= startDate.Date Then
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
        Else
            DateTimePicker1.CustomFormat = "dd MMMM yyyy"
            DateTimePicker2.CustomFormat = "dd MMMM yyyy"
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


    Private Sub DisableForm()

        For Each ctrl As Control In Controls
            If (Not ctrl.Equals(dgv) And (Not ctrl.Name = ProgressBar1.Name Or Not ctrl.Name = TextBox1.Name)) Then
                ctrl.Enabled = False
            End If
        Next

    End Sub

    Private Sub ShowCompletionDialog()                                                                                  ' Crea un'istanza del form modale e la mostra in modalità                                       

        Dim completedDownloadForm As New Form2()
        completedDownloadForm.ShowDialog()

    End Sub


    Private Sub ShowForm()

        Me.Show()

        If Me.WindowState = FormWindowState.Minimized Then
            Me.WindowState = FormWindowState.Normal
        End If

        Me.Activate()

    End Sub

    Private Sub reportCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        ShowForm()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Select Case ComboBox1.SelectedItem
            Case "E3"
                ComboBox3.SelectedIndex = 0
                Label6.Visible = True
                ComboBox3.Visible = True
            Case Else
                If ComboBox3.Visible Then
                    Label6.Visible = False
                    ComboBox3.Visible = False
                End If
        End Select

    End Sub

End Class

