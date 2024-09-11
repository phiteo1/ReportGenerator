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

    Enum State                  'State Machine of the downloading process
        DataLoading = 1
        TableLoading = 2
        SheetLoading = 3
        FinishedReport = 4
        Finished = 5
    End Enum

    Public Sub New()

        AddChimneyToList()

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

        Console.WriteLine(": " & Form1.startDate & Form1.endDate & Form1.section & Form1.reportType)

    End Sub
   

    Private Function GetFirstCaminiTable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, ByVal type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
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

    End Function



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

End Class
