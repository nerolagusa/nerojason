Imports System.IO
Imports System.Data.SqlClient
Imports ZXing
Public Class Savings
    Dim SqlConn As SqlConnection
    Dim COMMAND As SqlCommand

    Private Sub Savings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SqlConn = New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
        LoadData()

        PictureBox1.Size = New Size(145, 95)
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Check if ID and Deposit text boxes are filled
        If String.IsNullOrWhiteSpace(IDTextBox.Text) Then
            MessageBox.Show("Please fill up the ID textbox.")
            Return
        End If

        If String.IsNullOrWhiteSpace(Deposit.Text) Then
            MessageBox.Show("Please fill up the Deposit textbox.")
            Return
        End If

        ' Check if Deposit is a valid integer and does not exceed Int32.MaxValue
        Dim depositValue As Integer
        If Not Integer.TryParse(Deposit.Text, depositValue) OrElse depositValue < 0 Then
            MessageBox.Show("Please enter a valid positive deposit amount (integer).")
            Return
        End If

        If depositValue > Int32.MaxValue Then
            MessageBox.Show("Deposit amount cannot exceed the maximum limit of " & Int32.MaxValue & ".")
            Return
        End If

        ' Define the SQL command to update the Deposit, Total, and date columns of an existing record
        Dim updateCommand As String = "UPDATE dbo.Member SET Deposit = @Deposit, Total = Total + @Deposit, [date] = @Date WHERE ID = @ID"

        Try
            ' Open the connection
            SqlConn.Open()

            ' Create the SqlCommand object for updating
            COMMAND = New SqlCommand(updateCommand, SqlConn)

            ' Add parameters to prevent SQL injection
            COMMAND.Parameters.AddWithValue("@ID", IDTextBox.Text)
            COMMAND.Parameters.AddWithValue("@Deposit", depositValue)
            COMMAND.Parameters.AddWithValue("@Date", DateTimePicker1.Value) ' Add the date parameter

            ' Execute the update command
            Dim rowsAffected As Integer = COMMAND.ExecuteNonQuery()

            ' Check if any rows were updated
            If rowsAffected > 0 Then
                MessageBox.Show("Deposit amount and Total updated successfully!")

                ' Insert the deposit record into PaymentHistory
                Dim insertCommand As String = "INSERT INTO dbo.PaymentHistory (Withdraw, Deposit, PaymentHistoryID, [date], savings) VALUES (0, @Deposit, @PaymentHistoryID, @Date, 0)"

                ' Create the SqlCommand object for inserting
                Dim paymentHistoryID As Integer = Convert.ToInt32(IDTextBox.Text) ' Assuming PaymentHistoryID is the same as Member ID
                Dim insertCommandObj As New SqlCommand(insertCommand, SqlConn)

                ' Add parameters for the insert command
                insertCommandObj.Parameters.AddWithValue("@Deposit", depositValue)
                insertCommandObj.Parameters.AddWithValue("@PaymentHistoryID", paymentHistoryID)
                insertCommandObj.Parameters.AddWithValue("@Date", DateTimePicker1.Value)

                ' Execute the insert command
                insertCommandObj.ExecuteNonQuery()
            Else
                MessageBox.Show("No record found with the given ID.")
            End If
        Catch ex As Exception
            ' Handle any errors that may have occurred
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Always close the connection
            SqlConn.Close()
        End Try

        ' Clear the input fields and reset controls
        IDTextBox.Clear()
        Deposit.Clear()
        Widthdraw.Clear()
        Total.Clear()
        PictureBox1.Image = Nothing
        DateTimePicker1.Value = DateTime.Now
        DataGridView1.ClearSelection()
        DataGridView2.DataSource = Nothing
        TextBox2.Clear()

        ' Refresh the data in the DataGridView
        LoadData()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If String.IsNullOrWhiteSpace(IDTextBox.Text) Then
            MessageBox.Show("Please fill up the ID textbox.")
            Return
        End If

        If String.IsNullOrWhiteSpace(Widthdraw.Text) Then
            MessageBox.Show("Please fill up the Withdraw textbox.")
            Return
        End If

        Dim idValue As Integer
        If Not Integer.TryParse(IDTextBox.Text, idValue) Then
            MessageBox.Show("Please enter a valid ID.")
            Return
        End If

        ' Check if Withdraw is a valid integer
        Dim withdrawValue As Integer
        If Not Integer.TryParse(Widthdraw.Text, withdrawValue) Then
            MessageBox.Show("Please enter a valid withdrawal amount (integer).")
            Return
        End If

        ' Check the total amount from the database
        Dim totalAmount As Integer
        Dim totalCommand As String = "SELECT Total FROM dbo.Member WHERE ID = @ID"
        Try
            SqlConn.Open()
            Dim totalSqlCommand As New SqlCommand(totalCommand, SqlConn)
            totalSqlCommand.Parameters.AddWithValue("@ID", idValue)

            ' Retrieve the total amount for the given ID
            Dim result = totalSqlCommand.ExecuteScalar()
            If result IsNot Nothing Then
                totalAmount = Convert.ToInt32(result)

                ' Check if withdrawal exceeds total amount
                If withdrawValue > totalAmount Then
                    MessageBox.Show("Withdrawal amount cannot exceed the total amount.")
                    Return
                End If
            Else
                MessageBox.Show("No record found with the given ID.")
                Return
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving total amount: " & ex.Message)
            Return
        Finally
            SqlConn.Close()
        End Try

        Dim sqlCommand As String = "UPDATE dbo.Member SET Withdraw = @Withdraw, Total = Total - @Withdraw, [date] = @Date WHERE ID = @ID"
        Try
            SqlConn.Open()
            COMMAND = New SqlCommand(sqlCommand, SqlConn)
            COMMAND.Parameters.AddWithValue("@ID", idValue)
            COMMAND.Parameters.AddWithValue("@Withdraw", withdrawValue)
            COMMAND.Parameters.AddWithValue("@Date", DateTimePicker1.Value)

            Dim rowsAffected As Integer = COMMAND.ExecuteNonQuery()
            If rowsAffected > 0 Then
                MessageBox.Show("Withdrawal amount and Total updated successfully!")
            Else
                MessageBox.Show("No record found with the given ID.")
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            SqlConn.Close()
        End Try

        ' Insert into dbo.PaymentHistory with PaymentHistoryID as a foreign key
        Dim insertCommand As String = "INSERT INTO dbo.PaymentHistory (Withdraw, Deposit, PaymentHistoryID, [date]) VALUES (@Withdraw, @Deposit, @PaymentHistoryID, @Date)"
        Try
            SqlConn.Open()
            Dim insertSqlCommand As New SqlCommand(insertCommand, SqlConn)

            ' Assuming Deposit is zero for withdrawal
            insertSqlCommand.Parameters.AddWithValue("@Withdraw", withdrawValue)
            insertSqlCommand.Parameters.AddWithValue("@Deposit", 0) ' Adjust this if needed
            insertSqlCommand.Parameters.AddWithValue("@PaymentHistoryID", idValue) ' Use the foreign key from dbo.Member
            insertSqlCommand.Parameters.AddWithValue("@Date", DateTimePicker1.Value)

            insertSqlCommand.ExecuteNonQuery()
            MessageBox.Show("Payment history recorded successfully!")
        Catch ex As Exception
            MessageBox.Show("An error occurred while inserting into payment history: " & ex.Message)
        Finally
            SqlConn.Close()
        End Try

        ' Clear inputs and refresh
        IDTextBox.Clear()
        Deposit.Clear()
        Widthdraw.Clear()
        Total.Clear()
        PictureBox1.Image = Nothing
        DateTimePicker1.Value = DateTime.Now
        DataGridView1.ClearSelection()
        DataGridView2.DataSource = Nothing
        TextBox2.Clear()

        ' Refresh the data in the DataGridView
        LoadData()
    End Sub
    Private Sub Widthdraw_TextChanged(sender As Object, e As EventArgs) Handles Widthdraw.TextChanged

    End Sub
    Private Sub Widthdraw_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Widthdraw.KeyPress
        ' Check if the pressed key is not a digit and not a control key (like Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Suppress the key press event (do not allow input)
            e.Handled = True
        End If
    End Sub

    Private Sub IDTextBox_TextChanged(sender As Object, e As EventArgs) Handles IDTextBox.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        ' Ensure the click is on a valid row and not the header
        If e.RowIndex >= 0 Then
            ' Get the selected row
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Populate the controls with the selected row data
            IDTextBox.Text = row.Cells("ID").Value.ToString() ' ID
            TextBox2.Text = row.Cells("Name").Value.ToString() ' Name
            DateTimePicker1.Value = Convert.ToDateTime(row.Cells("date").Value) ' Date

            ' Load the picture from the database
            If Not IsDBNull(row.Cells("Picture").Value) Then
                Dim imageData As Byte() = CType(row.Cells("Picture").Value, Byte())
                Using ms As New MemoryStream(imageData)
                    PictureBox1.Image = Image.FromStream(ms)
                End Using
            Else
                PictureBox1.Image = Nothing ' Set to no image if picture is null
            End If

            ' Load the QR code from the database
            If Not IsDBNull(row.Cells("qrcode").Value) Then
                Dim qrCodeData As Byte() = CType(row.Cells("qrcode").Value, Byte())
                Using ms As New MemoryStream(qrCodeData)
                    qrcodepic.Image = Image.FromStream(ms) ' Display QR code in PictureBox
                End Using
            Else
                qrcodepic.Image = Nothing ' Set to no image if QR code is null
            End If

            ' Set the sizes and modes of the PictureBoxes
            PictureBox1.Size = New Size(145, 95)
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

            qrcodepic.Size = New Size(141, 88)
            qrcodepic.SizeMode = PictureBoxSizeMode.StretchImage

            ' Load data for DataGridView2 matching the selected ID
            Dim selectedID As Integer
            If Integer.TryParse(IDTextBox.Text, selectedID) Then
                ' Clear DataGridView2 before loading new data
                DataGridView2.DataSource = Nothing
                DataGridView2.ClearSelection()
                LoadDataForDataGridView2(selectedID) ' Load relevant data for DataGridView2
            End If
        End If
    End Sub

    Private Sub Deposit_TextChanged(sender As Object, e As EventArgs) Handles Deposit.TextChanged

    End Sub
    Private Sub Deposit_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Deposit.KeyPress
        ' Check if the pressed key is not a digit and not a control key (like Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Suppress the key press event (do not allow input)
            e.Handled = True
        End If
    End Sub
    Private originalDataTable As DataTable

    Private Sub LoadData()
        Dim sqlCommand As String = "SELECT ID, Name, [date], Picture, qrcode, Withdraw, Deposit, Total FROM dbo.Member"

        Try
            ' Open the connection
            SqlConn.Open()

            ' Create a new SqlDataAdapter
            Dim adapter As New SqlDataAdapter(sqlCommand, SqlConn)

            ' Create a new DataTable
            originalDataTable = New DataTable() ' Initialize the original DataTable

            ' Fill the DataTable with data from the database
            adapter.Fill(originalDataTable)

            ' Bind the DataTable to the DataGridView
            DataGridView1.DataSource = originalDataTable

            ' Configure the DataGridView to show images
            ConfigureDataGridView()
        Catch ex As Exception
            ' Handle any errors that may have occurred
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Always close the connection
            SqlConn.Close()
        End Try

    End Sub
    Private Sub SearchByID()
        ' Get the search text from TextBox1
        Dim searchText As String = TextBox1.Text.Trim()

        ' Check if the original DataTable is not null
        If originalDataTable IsNot Nothing Then
            ' Create a DataView to apply the filter
            Dim dataView As New DataView(originalDataTable)

            ' Apply the filter based only on the search text for ID
            If String.IsNullOrEmpty(searchText) Then
                dataView.RowFilter = String.Empty ' Show all rows if search text is empty
            Else
                ' Make sure the search text is numeric since ID is an integer
                Dim isNumeric As Boolean = Integer.TryParse(searchText, Nothing)

                If isNumeric Then
                    dataView.RowFilter = $"[ID] = {searchText}" ' Filter by exact match for ID
                Else
                    MessageBox.Show("Please enter a valid numeric ID.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            End If

            ' Update the DataGridView with the filtered data
            DataGridView1.DataSource = dataView
        End If
    End Sub
    Private Sub ConfigureDataGridView()
        ' Set the size of the DataGridView
        DataGridView1.Size = New Size(530, 171)

        ' Configure the DataGridView to display images
        For Each column As DataGridViewColumn In DataGridView1.Columns
            If column.Name = "Picture" Then
                column.HeaderText = "Picture"
                column.CellTemplate = New DataGridViewImageCell()
            ElseIf column.Name = "qrcode" Then
                column.HeaderText = "QR Code"
                column.CellTemplate = New DataGridViewImageCell()
            ElseIf column.Name = "ID" Then
                column.HeaderText = "ID"
            ElseIf column.Name = "Name" Then
                column.HeaderText = "Name"
            ElseIf column.Name = "Total" Then
                column.HeaderText = "Total"
            ElseIf column.Name = "date" Then
                column.Visible = False ' Hide the date column
            ElseIf column.Name = "Withdraw" OrElse column.Name = "Deposit" Then
                column.Visible = False ' Hide the Withdraw and Deposit columns
            End If
        Next

        ' Auto-size columns based on content
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1.AllowUserToAddRows = False
    End Sub
    Private Sub Total_TextChanged(sender As Object, e As EventArgs) Handles Total.TextChanged

    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            IDTextBox.Text = row.Cells("ID").Value.ToString()
            TextBox2.Text = row.Cells("Name").Value.ToString()
            Total.Text = row.Cells("Total").Value.ToString()

            If row.Cells("Picture").Value IsNot Nothing Then
                Dim pictureData As Byte() = CType(row.Cells("Picture").Value, Byte())
                Using ms As New MemoryStream(pictureData)
                    PictureBox1.Image = Image.FromStream(ms)
                End Using
            Else
                PictureBox1.Image = Nothing
            End If

            ' Ensure PictureBox size

        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        LoadData()

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub
    Private Sub LoadDataForDataGridView2(selectedID As Integer)
        ' Use selectedID to load data based on the specific ID
        Dim sqlCommand As String = "SELECT [date], Withdraw, Deposit FROM dbo.PaymentHistory WHERE PaymentHistoryID = @ID"

        Try
            ' Open the connection
            SqlConn.Open()

            ' Create a new SqlCommand with parameter
            Using command As New SqlCommand(sqlCommand, SqlConn)
                command.Parameters.AddWithValue("@ID", selectedID) ' Use selectedID here

                ' Create a new SqlDataAdapter
                Dim adapter As New SqlDataAdapter(command)

                ' Create a new DataTable
                Dim dataTable As New DataTable()

                ' Fill the DataTable with data from the database
                adapter.Fill(dataTable)

                ' Bind the DataTable to the DataGridView
                DataGridView2.DataSource = dataTable

                ' Configure the DataGridView
                DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                DataGridView2.AllowUserToAddRows = False
                DataGridView2.Columns("date").HeaderText = "Date"
                DataGridView2.Columns("Withdraw").HeaderText = "Withdraw"
                DataGridView2.Columns("Deposit").HeaderText = "Deposit"
            End Using
        Catch ex As Exception
            ' Handle any errors that may have occurred
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Always close the connection
            SqlConn.Close()
        End Try
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        IDTextBox.Clear()
        Deposit.Clear()
        Widthdraw.Clear()
        Total.Clear()

        '
        PictureBox1.Image = Nothing


        DateTimePicker1.Value = DateTime.Now


        DataGridView1.ClearSelection()
        DataGridView2.DataSource = Nothing


        TextBox2.Clear()


        LoadData()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim searchText As String = TextBox1.Text.Trim()

        ' Ensure that the DataGridView is bound to the original DataTable
        If originalDataTable IsNot Nothing Then
            ' Create a DataView to apply the filter
            Dim dataView As New DataView(originalDataTable)

            ' Apply the filter based only on the search text for ID
            If String.IsNullOrEmpty(searchText) Then
                dataView.RowFilter = String.Empty ' Show all rows if search text is empty
            Else
                ' Make sure the search text is numeric since ID is an integer
                Dim isNumeric As Boolean = Integer.TryParse(searchText, Nothing)

                If isNumeric Then
                    dataView.RowFilter = $"[ID] = {searchText}" ' Filter by exact match for ID
                Else
                    MessageBox.Show("Please enter a valid numeric ID.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            End If

            ' Update the DataGridView with the filtered data
            DataGridView1.DataSource = dataView
        End If

    End Sub


    Private Sub HomeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HomeToolStripMenuItem.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub awit_Click(sender As Object, e As EventArgs) Handles awit.Click
        Me.Hide()
        member.Show()
    End Sub

    Private Sub ContributionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContributionToolStripMenuItem.Click
        Me.Hide()
        Contribution.Show()

    End Sub

    Private Sub LoanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoanToolStripMenuItem.Click
        Me.Hide()
        Loan.Show()
    End Sub


    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        ' Refresh the data in DataGridView1 and DataGridView2
        LoadData()

        ' Clear any selections in DataGridView2
        DataGridView2.DataSource = Nothing
        DataGridView2.ClearSelection()

        ' Optionally clear any relevant fields like text boxes, images, etc.
        IDTextBox.Clear()
        TextBox2.Clear()
        Deposit.Clear()
        Widthdraw.Clear()
        Total.Clear()
        PictureBox1.Image = Nothing
        qrcodepic.Image = Nothing
        DateTimePicker1.Value = DateTime.Now

        MessageBox.Show("Data refreshed successfully!")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Open the SQL connection if it's closed
        If SqlConn.State = ConnectionState.Closed Then
            SqlConn.Open()
        End If

        Try
            ' Retrieve the ID from TextBox4
            Dim idValue As Integer
            If Not Integer.TryParse(IDTextBox.Text, idValue) Then

                Exit Sub
            End If

            ' SQL Query to fetch data for the specified ID
            Dim query As String = "SELECT Name, Total, Contribution, Picture,qrcode FROM dbo.member WHERE ID = @ID"

            Using COMMAND As New SqlCommand(query, SqlConn)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = COMMAND.ExecuteReader()
                    If reader.Read() Then
                        ' Populate the TextBoxes and Label with retrieved values
                        TextBox2.Text = reader("Name").ToString()                     ' Name
                        Total.Text = reader("Total").ToString()                    ' Total


                        ' Retrieve the image from the database
                        If Not IsDBNull(reader("Picture")) Then
                            Dim imgData As Byte() = CType(reader("Picture"), Byte())
                            Using ms As New System.IO.MemoryStream(imgData)
                                PictureBox1.Image = Image.FromStream(ms)
                            End Using
                            ' Set the PictureBox SizeMode to StretchImage
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                        Else
                            PictureBox1.Image = Nothing
                        End If

                        If Not IsDBNull(reader("qrcode")) Then
                            Dim imgData As Byte() = CType(reader("qrcode"), Byte())
                            Using ms As New System.IO.MemoryStream(imgData)
                                qrcodepic.Image = Image.FromStream(ms)
                            End Using
                            ' Set the PictureBox SizeMode to StretchImage
                            qrcodepic.SizeMode = PictureBoxSizeMode.StretchImage
                        Else
                            qrcodepic.Image = Nothing
                        End If


                    Else
                        MessageBox.Show("No record found for the provided ID.")



                    End If

                End Using
            End Using

            Dim queryForDataGrid As String = "SELECT date,withdraw,deposit FROM dbo.PaymentHistory WHERE PaymentHistoryID = @ID"

            Using COMMAND As New SqlCommand(queryForDataGrid, SqlConn)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and load the results into a DataTable
                Dim dataAdapter As New SqlDataAdapter(COMMAND)
                Dim dataTable As New DataTable()
                dataAdapter.Fill(dataTable)

                ' Bind the result to DataGridView2
                DataGridView2.DataSource = dataTable
                DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            End Using

        Catch ex As Exception
            ' Handle any errors that may have occurred
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Close the connection if it's open
            If SqlConn.State = ConnectionState.Open Then
                SqlConn.Close()
            End If
        End Try
    End Sub

    Private Sub qrcodepic_Click(sender As Object, e As EventArgs) Handles qrcodepic.Click

    End Sub
End Class