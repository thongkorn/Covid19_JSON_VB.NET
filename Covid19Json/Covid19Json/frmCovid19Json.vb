#Region "ABOUT"
' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Coronavirus API for Current cases by country COVID-19.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------
#End Region

Imports System.Net
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Public Class frmCovid19Json

    Dim table As New DataTable

    Private Sub frmCovid19Json_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call GetCountry()
    End Sub

    Private Sub cmbCountry_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbCountry.SelectedIndexChanged
        Dim url As String = String.Empty
        Select Case cmbCountry.Text
            Case "- All Countries Info -"
                url = "https://coronavirus-19-api.herokuapp.com/countries"
            Case Else
                url = "https://coronavirus-19-api.herokuapp.com/countries/" & cmbCountry.Text
        End Select
        '//
        Call GetDataJson(url)
        'Call GetDataJsonObject(url)
    End Sub

    '// Deserialized JSON and return DataTable.
    Public Function DeserializeDataTable(json As String) As DataTable
        Dim dt As DataTable = TryCast(JsonConvert.DeserializeObject(json, (GetType(DataTable))), DataTable)
        Return dt
    End Function

    Private Sub GetCountry()
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Dim url As String = "https://coronavirus-19-api.herokuapp.com/countries"
        Try
            request = DirectCast(WebRequest.Create(url), HttpWebRequest)
            response = DirectCast(request.GetResponse(), HttpWebResponse)
            reader = New StreamReader(response.GetResponseStream())
            '//
            Dim s As String
            s = reader.ReadToEnd
            '// Replace null value from JSON to zero.
            s = s.Replace("null", "0")
            Dim dt As DataTable
            dt = DeserializeDataTable(s)
            With cmbCountry
                .Items.Add("- All Countries Info -")
            End With
            table.Columns.Add("Country", GetType(String))
            '// Add all country into ComboBox.
            For i = 0 To dt.Rows.Count - 1
                cmbCountry.Items.Add(dt.Rows(i).Item(0).ToString)
                '// AutoComplete.
                table.Rows.Add(dt.Rows(i).Item(0).ToString)
            Next
            cmbCountry.SelectedIndex = 0
            cmbCountry.Sorted = True

            '// AutoComplete in TextBox Control.
            Dim DataCollection As New AutoCompleteStringCollection()
            DataCollection = GetAutoSourceCollection(table)
            txtSearch.AutoCompleteCustomSource = DataCollection
            txtSearch.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            txtSearch.AutoCompleteSource = AutoCompleteSource.CustomSource

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    '// AutoComplete Collection
    Private Function GetAutoSourceCollection(ByVal table As DataTable) As AutoCompleteStringCollection
        Dim AutoSourceCollection As AutoCompleteStringCollection = New AutoCompleteStringCollection()
        '// table from countries.
        For Each row As DataRow In table.Rows
            AutoSourceCollection.Add(row(0).ToString) '// Country data is the first column.
        Next
        Return AutoSourceCollection
    End Function

    '// Get JSON From URL.
    Private Sub GetDataJson(ByVal url As String)
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Try
            request = DirectCast(WebRequest.Create(url), HttpWebRequest)
            response = DirectCast(request.GetResponse(), HttpWebResponse)
            reader = New StreamReader(response.GetResponseStream())
            Dim s As String
            s = reader.ReadToEnd
            s = s.Replace("null", "0")
            Dim dt As DataTable
            If Microsoft.VisualBasic.Left(s, 1) = "[" Then
                dt = DeserializeDataTable(s)
            Else
                dt = DeserializeDataTable("[" & s & "]")
            End If
            dgvData.DataSource = dt
            Call SetupGridView()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    '// Get JSON by Object From URL.
    Private Sub GetDataJsonObject(ByVal url As String)
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Try
            request = DirectCast(WebRequest.Create(url), HttpWebRequest)
            response = DirectCast(request.GetResponse(), HttpWebResponse)
            reader = New StreamReader(response.GetResponseStream())
            Dim s As String
            s = reader.ReadToEnd.Replace("null", "0")
            If Microsoft.VisualBasic.Left(s, 1) <> "[" Then
                s = "[" & s & "]"
            End If
            '//
            Dim res() = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ItemsCovid())(s)
            Dim dt As New DataTable
            dt.Columns.Add("country", GetType(String))
            dt.Columns.Add("cases")
            dt.Columns.Add("todayCases")
            dt.Columns.Add("deaths")
            dt.Columns.Add("todayDeaths")
            dt.Columns.Add("recovered")
            dt.Columns.Add("active")
            dt.Columns.Add("critical")
            dt.Columns.Add("casesPerOneMillion", GetType(Integer))
            For Each covid As ItemsCovid In res
                Dim dr As DataRow = dt.NewRow()
                If covid IsNot Nothing Then
                    dr(0) = covid.country
                    dr(1) = covid.cases
                    dr(2) = covid.todayCases
                    dr(3) = covid.deaths
                    dr(4) = covid.todayDeaths
                    dr(5) = covid.recovered
                    dr(6) = covid.active
                    dr(7) = covid.critical
                    dr(8) = covid.casesPerOneMillion
                End If
                '// Add row.
                dt.Rows.Add(dr)
            Next
            dgvData.DataSource = dt
            Call SetupGridView()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtSearch_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        '// Press Enter
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            '// Search for Country in DataTable.
            Dim dv As DataView = New DataView(table)
            dv.RowFilter = "Country = " & "'" & txtSearch.Text & "'"
            If dv.Count > 0 Then
                Dim url As String = "https://coronavirus-19-api.herokuapp.com/countries/" & txtSearch.Text
                '// Second Methods
                Call GetDataJsonObject(url)
            End If
        End If
    End Sub

    Private Sub SetupGridView()
        With dgvData
            .RowHeadersVisible = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .ReadOnly = True
            '// Data rows
            .Font = New Font("Tahoma", 9)
            .RowTemplate.MinimumHeight = 27
            .RowTemplate.Height = 27
            '// Column Header
            .ColumnHeadersHeight = 30
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            '// Autosize Column
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '// Header
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.White
                .Font = New Font(dgvData.Font, FontStyle.Bold)
            End With
        End With
        '/ todayDeaths
        For i = 0 To dgvData.RowCount - 1
            If dgvData.Rows(i).Cells(4).Value > 0 Then
                dgvData.Rows(i).Cells(4).Style.BackColor = Color.Red
            End If
        Next
        For i = 1 To dgvData.ColumnCount - 1
            dgvData.Columns(i).DefaultCellStyle.Format = "N0"
            With dgvData.Columns(i)
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With
        Next
    End Sub

    Private Sub frmCovid19Json_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class

'// "country":"Thailand","cases":599,"todayCases":188,"deaths":1,"todayDeaths":0,"recovered":44,"active":554,"critical":7,"casesPerOneMillion":9
Public Class ItemsCovid
    Public Property country As String
    Public Property cases As Integer
    Public Property todayCases As Integer
    Public Property deaths As Integer
    Public Property todayDeaths As Integer
    Public Property recovered As Integer
    Public Property active As Integer
    Public Property critical As Integer
    Public Property casesPerOneMillion As Integer
End Class
