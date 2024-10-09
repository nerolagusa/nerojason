Imports System.Data.SqlClient

Public Class PaymentHistoryForm
    Private connectionString As String = "Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1"

    ' This property will receive the member ID from the main form
    Public Property MemberId As String

    Private Sub PaymentHistoryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadPaymentHistory(MemberId)
    End Sub

    Public Sub LoadPaymentHistory(memberId As String)
        ' Update the query to select data from the PaymentHistory table
        Dim query As String = "SELECT [Payment], [date] FROM dbo.PaymentHistory WHERE [PaymentHistoryID] = @ID"

        Using connection As New SqlConnection(connectionString)
            Dim adapter As New SqlDataAdapter(query, connection)
            Dim dataTable As New DataTable()

            ' Add parameter to prevent SQL injection
            adapter.SelectCommand.Parameters.AddWithValue("@ID", memberId)

            Try
                connection.Open()
                adapter.Fill(dataTable)

                ' Set the DataGridView DataSource
                DataGridView1.DataSource = dataTable

                ' Optionally, configure DataGridView properties
                DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                DataGridView1.AllowUserToAddRows = False
            Catch ex As SqlException
                MessageBox.Show("An error occurred while retrieving payment history: " & ex.Message)
            Catch ex As Exception
                MessageBox.Show("An error occurred: " & ex.Message)
            Finally
                connection.Close()
            End Try
        End Using
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        ' You can handle cell clicks if needed
    End Sub
End Class
