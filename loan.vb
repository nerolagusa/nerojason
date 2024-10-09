Imports System.Data.SqlClient

Public Class Loan
    Dim SqlConn As SqlConnection
    Dim COMMAND As SqlCommand

    Private Sub Loan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize the connection string for the SQL Server
        SqlConn = New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
        DataGridView1.Width = 705
        DataGridView1.Height = 151
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

        ' Load data into the DataGridView
        LoadData()
    End Sub



    ' Method to load data into the DataGridView
    Private Sub LoadData()
        Dim connectionString As String = "Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1"
        Dim query As String = "SELECT [ID], [name], [interest], [date], [amount], [payment], [monthly_terms], [PaymentMonths] FROM dbo.Member" ' Include PaymentMonths

        Using connection As New SqlConnection(connectionString)
            Dim adapter As New SqlDataAdapter(query, connection)
            Dim dataTable As New DataTable()

            Try
                connection.Open()
                adapter.Fill(dataTable)
            Catch ex As SqlException
                MessageBox.Show("An error occurred while retrieving data: " & ex.Message)
            Catch ex As Exception
                MessageBox.Show("An error occurred: " & ex.Message)
            Finally
                connection.Close()
            End Try

            ' Set AllowUserToAddRows to False to remove the empty row
            DataGridView1.AllowUserToAddRows = False
            DataGridView1.DataSource = dataTable

            ' Hide the 'date' and 'payment' columns if needed
            DataGridView1.Columns("date").Visible = False
            DataGridView1.Columns("payment").Visible = False

            ' Ensure that the Monthly_terms column is displayed (it should be)
            ' Optionally, you can also hide it if you don't want it visible
            DataGridView1.Columns("monthly_terms").Visible = True
        End Using
    End Sub

    ' Event handler for DataGridView cell click
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Check if the click is on a valid cell (not the header)
        If e.RowIndex >= 0 Then
            ' Get the selected row
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Populate TextBoxes with selected row data
            IDTextBox.Text = row.Cells("ID").Value.ToString()
            NameTextBox.Text = row.Cells("name").Value.ToString()
            DatePicker.Value = Convert.ToDateTime(row.Cells("date").Value)
            AmountTextBox.Text = row.Cells("amount").Value.ToString()
            InterestTextBox.Text = row.Cells("interest").Value.ToString()
            TextBox1.Text = row.Cells("monthly_terms").Value.ToString()

            TextBox2.Text = row.Cells("PaymentMonths").Value.ToString() ' Assuming you have a PaymentMonthsTextBox
        End If
    End Sub

    Private Sub update_Click(sender As Object, e As EventArgs) Handles update.Click
        ' Get the values from the TextBoxes and DateTimePicker
        Dim idText As String = IDTextBox.Text
        Dim interestValue As Double
        Dim amountValue As Double
        Dim paymentValue As Double = 0 ' Set default payment value to 0
        Dim dateValue As Date = DatePicker.Value
        Dim monthlyTermsValue As Integer ' Variable to hold the monthly terms

        ' Validate and convert interest and amount to Double
        If Not Double.TryParse(InterestTextBox.Text, interestValue) Then
            MessageBox.Show("Invalid interest value. Please enter a valid number.")
            Exit Sub
        End If

        If Not Double.TryParse(AmountTextBox.Text, amountValue) Then
            MessageBox.Show("Invalid amount entered. Please enter a valid number.")
            Exit Sub
        End If

        ' If PaymentTextBox is not filled, it will use the default paymentValue (0)
        If Not String.IsNullOrWhiteSpace(PaymentTextBox.Text) AndAlso Not Double.TryParse(PaymentTextBox.Text, paymentValue) Then
            MessageBox.Show("Invalid payment entered. Please enter a valid number.")
            Exit Sub
        End If

        ' Validate and convert monthly terms to Integer
        If Not Integer.TryParse(TextBox1.Text, monthlyTermsValue) Then
            MessageBox.Show("Invalid monthly terms entered. Please enter a valid integer.")
            Exit Sub
        End If

        ' Calculate the interest amount (convert percentage to decimal by dividing by 100)
        Dim interestAmount As Double = amountValue * (interestValue / 100)

        ' Add the interest to the original amount
        Dim newAmount As Double = amountValue + interestAmount

        ' Update the AmountTextBox with the new calculated amount
        AmountTextBox.Text = newAmount.ToString("F2") ' Format to 2 decimal places

        Try
            ' Open the SQL connection
            SqlConn.Open()

            ' SQL Update Query to update an existing record in the Member table
            Dim updateQuery As String = "UPDATE dbo.Member SET [name] = @name, [interest] = @interest, [date] = @date, [amount] = @amount, [payment] = @payment, [monthly_terms] = @monthly_terms WHERE [ID] = @id"
            Using command As New SqlCommand(updateQuery, SqlConn)
                ' Add parameters to prevent SQL injection
                command.Parameters.AddWithValue("@id", idText)
                command.Parameters.AddWithValue("@name", NameTextBox.Text) ' Update the name if needed
                command.Parameters.AddWithValue("@interest", interestValue)
                command.Parameters.AddWithValue("@date", dateValue)
                command.Parameters.AddWithValue("@amount", newAmount) ' Use the calculated new amount
                command.Parameters.AddWithValue("@payment", paymentValue) ' Use the payment value
                command.Parameters.AddWithValue("@monthly_terms", monthlyTermsValue) ' Add monthly terms

                ' Execute the query
                Dim rowsAffected As Integer = command.ExecuteNonQuery()

                ' Check if the update was successful
                If rowsAffected > 0 Then
                    MessageBox.Show("Loan record updated successfully!")
                    ' Optionally reload data into the DataGridView
                    LoadData()
                Else
                    MessageBox.Show("Failed to update the loan record. Ensure the record with the specified ID exists.")
                End If
            End Using
        Catch ex As SqlException
            ' Handle SQL exception
            MessageBox.Show("A database error occurred: " & ex.Message)
        Catch ex As Exception
            ' Handle general exception
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Close the connection if it was opened
            If SqlConn.State = ConnectionState.Open Then
                SqlConn.Close()
            End If
        End Try
    End Sub

    Private Sub Delete_Click(sender As Object, e As EventArgs) Handles Delete.Click
        ' Get the value from the IDTextBox
        Dim idText As String = IDTextBox.Text

        ' Validate that an ID is entered
        If String.IsNullOrEmpty(idText) Then
            MessageBox.Show("Please enter an ID to delete.")
            Exit Sub
        End If

        ' Ask for confirmation before deleting
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete the record with ID '" & idText & "'?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' Check if the user clicked Yes
        If result = DialogResult.Yes Then
            Try
                ' Open the SQL connection
                SqlConn.Open()

                ' Begin a transaction
                Using transaction As SqlTransaction = SqlConn.BeginTransaction()
                    ' SQL Delete Query to remove a record from the Member table based on the ID
                    Dim deleteMemberQuery As String = "DELETE FROM dbo.Member WHERE [ID] = @id"
                    Using memberCommand As New SqlCommand(deleteMemberQuery, SqlConn, transaction)
                        ' Add parameter to prevent SQL injection
                        memberCommand.Parameters.AddWithValue("@id", idText)

                        ' Execute the query
                        Dim rowsAffected As Integer = memberCommand.ExecuteNonQuery()

                        ' Check if the deletion was successful
                        If rowsAffected > 0 Then
                            ' SQL Delete Query to remove corresponding records from PaymentHistory
                            Dim deletePaymentHistoryQuery As String = "DELETE FROM dbo.PaymentHistory WHERE PaymentHistoryID = @id"
                            Using paymentHistoryCommand As New SqlCommand(deletePaymentHistoryQuery, SqlConn, transaction)
                                ' Add parameter to prevent SQL injection
                                paymentHistoryCommand.Parameters.AddWithValue("@id", idText)

                                ' Execute the query for deleting payment history
                                Dim paymentHistoryRowsAffected As Integer = paymentHistoryCommand.ExecuteNonQuery()

                                ' Commit the transaction
                                transaction.Commit()
                                MessageBox.Show("Record and associated payment history deleted successfully!")

                                ' Reload data into the DataGridView or update UI if needed
                                LoadData()
                            End Using
                        Else
                            MessageBox.Show("No loan record found with the provided ID.")
                        End If
                    End Using
                End Using
            Catch ex As SqlException
                ' Handle SQL exception
                MessageBox.Show("A database error occurred: " & ex.Message)
            Catch ex As Exception
                ' Handle general exception
                MessageBox.Show("An error occurred: " & ex.Message)
            Finally
                ' Close the connection if it was opened
                If SqlConn.State = ConnectionState.Open Then
                    SqlConn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub FilterData(searchText As String)
        Dim connectionString As String = "Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1"
        Dim query As String = "SELECT [ID], [name], [interest], [date], [amount], [payment] FROM dbo.Member WHERE [name] LIKE @searchText"

        Using connection As New SqlConnection(connectionString)
            Dim adapter As New SqlDataAdapter(query, connection)
            Dim dataTable As New DataTable()

            ' Add parameter to prevent SQL injection
            adapter.SelectCommand.Parameters.AddWithValue("@searchText", "%" & searchText & "%")

            Try
                connection.Open()
                adapter.Fill(dataTable)
            Catch ex As SqlException
                MessageBox.Show("An error occurred while retrieving data: " & ex.Message)
            Catch ex As Exception
                MessageBox.Show("An error occurred: " & ex.Message)
            Finally
                connection.Close()
            End Try

            ' Update DataGridView with filtered data
            DataGridView1.DataSource = dataTable
        End Using
    End Sub

    Private Sub NameTextBox_TextChanged(sender As Object, e As EventArgs) Handles NameTextBox.TextChanged
        ' Get the current text in the TextBox
        Dim text As String = NameTextBox.Text

        ' Remove numeric characters from the text
        Dim newText As String = String.Concat(text.Where(Function(c) Not Char.IsDigit(c)))

        ' Update the TextBox with the filtered text
        If text <> newText Then
            ' Temporarily remove the event handler to avoid recursive calls
            RemoveHandler NameTextBox.TextChanged, AddressOf NameTextBox_TextChanged
            NameTextBox.Text = newText
            ' Restore the event handler
            AddHandler NameTextBox.TextChanged, AddressOf NameTextBox_TextChanged

            ' Optionally, move the cursor to the end of the TextBox
            NameTextBox.SelectionStart = NameTextBox.Text.Length
        End If
    End Sub

    Private Sub PaymentTextBox_TextChanged(sender As Object, e As EventArgs) Handles PaymentTextBox.TextChanged
        Dim newText As String = PaymentTextBox.Text

        ' Use a StringBuilder to create a new string with only valid characters
        Dim validText As New System.Text.StringBuilder()

        Dim decimalAdded As Boolean = False ' To track if a decimal point has been added

        For Each c As Char In newText
            ' Check if the character is a digit, a negative sign (at the beginning), or a decimal point (only one allowed)
            If Char.IsDigit(c) Then
                validText.Append(c)
            ElseIf c = "-"c AndAlso validText.Length = 0 Then
                validText.Append(c) ' Allow negative sign, but only at the start
            ElseIf c = "."c AndAlso Not decimalAdded Then
                validText.Append(c) ' Allow only one decimal point
                decimalAdded = True
            End If
        Next

        ' Update the TextBox with the valid characters only
        PaymentTextBox.Text = validText.ToString()

        ' Set the cursor position to the end of the TextBox
        PaymentTextBox.SelectionStart = PaymentTextBox.Text.Length
    End Sub

    Private Sub InterestTextBox_TextChanged(sender As Object, e As EventArgs) Handles InterestTextBox.TextChanged

        Dim text As String = InterestTextBox.Text

        ' Remove non-numeric characters from the text
        Dim newText As String = String.Concat(text.Where(Function(c) Char.IsDigit(c)))

        ' Update the TextBox with the filtered text
        If text <> newText Then
            ' Temporarily remove the event handler to avoid recursive calls
            RemoveHandler InterestTextBox.TextChanged, AddressOf InterestTextBox_TextChanged
            InterestTextBox.Text = newText
            ' Restore the event handler
            AddHandler InterestTextBox.TextChanged, AddressOf InterestTextBox_TextChanged

            ' Optionally, move the cursor to the end of the TextBox
            InterestTextBox.SelectionStart = InterestTextBox.Text.Length
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub AmountTextBox_TextChanged(sender As Object, e As EventArgs) Handles AmountTextBox.TextChanged
        Dim newText As String = AmountTextBox.Text

        ' Use a StringBuilder to create a new string with only valid characters
        Dim validText As New System.Text.StringBuilder()

        Dim decimalAdded As Boolean = False ' To track if a decimal point has been added

        For Each c As Char In newText
            ' Check if the character is a digit, a negative sign (at the beginning), or a decimal point (only one allowed)
            If Char.IsDigit(c) Then
                validText.Append(c)
            ElseIf c = "-"c AndAlso validText.Length = 0 Then
                validText.Append(c) ' Allow negative sign, but only at the start
            ElseIf c = "."c AndAlso Not decimalAdded Then
                validText.Append(c) ' Allow only one decimal point
                decimalAdded = True
            End If
        Next

        ' Update the TextBox with the valid characters only
        AmountTextBox.Text = validText.ToString()

        ' Set the cursor position to the end of the TextBox
        AmountTextBox.SelectionStart = AmountTextBox.Text.Length
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Retrieve the values from the TextBoxes
        Dim amountValue As Double
        Dim paymentValue As Double
        Dim paymentDate As Date = DatePicker.Value ' Get the selected date from the DatePicker
        Dim monthlyTerms As Integer
        Dim paymentMonths As Integer ' Variable to hold the current PaymentMonths

        ' Validate and convert amount and payment to Double
        If Not Double.TryParse(AmountTextBox.Text, amountValue) Then
            MessageBox.Show("Invalid amount entered. Please enter a valid number.")
            Exit Sub
        End If

        If Not Double.TryParse(PaymentTextBox.Text, paymentValue) Then
            MessageBox.Show("Invalid payment entered. Please enter a valid number.")
            Exit Sub
        End If

        ' Calculate the new amount after subtracting the payment
        Dim updatedAmount As Double = amountValue - paymentValue
        Dim change As Double = 0 ' Variable to store change

        ' Set the amount to zero if payment exceeds the available amount
        If updatedAmount < 0 Then
            change = Math.Abs(updatedAmount) ' Calculate change
            updatedAmount = 0
            MessageBox.Show($"Payment Successful. Payment cleared. Change to be returned: {change.ToString("F2")}")
        Else
            ' Show the remaining amount if payment doesn't exceed
            MessageBox.Show($"Payment Successful. Remaining amount: {updatedAmount.ToString("F2")}")
        End If

        ' Update the AmountTextBox with the new calculated amount
        AmountTextBox.Text = updatedAmount.ToString("F2")

        ' Fetch current PaymentMonths from the database
        Try
            SqlConn.Open()
            Dim paymentMonthsQuery As String = "SELECT [PaymentMonths] FROM dbo.Member WHERE [ID] = @id"
            Using paymentMonthsCommand As New SqlCommand(paymentMonthsQuery, SqlConn)
                paymentMonthsCommand.Parameters.AddWithValue("@id", IDTextBox.Text)
                paymentMonths = Convert.ToInt32(paymentMonthsCommand.ExecuteScalar()) ' Get the current PaymentMonths
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while fetching PaymentMonths: " & ex.Message)
            Return
        Finally
            If SqlConn.State = ConnectionState.Open Then
                SqlConn.Close()
            End If
        End Try


        Dim paymentDeduction As Integer = CInt(paymentValue)

        ' Deduct payment from PaymentMonths, but ensure it doesn't go below zero
        If paymentMonths - paymentDeduction < 0 Then
            paymentMonths = 0
        Else
            paymentMonths -= paymentDeduction
        End If

        ' Check if monthlyTerms should be reduced based on PaymentMonths or remaining amount
        If Integer.TryParse(TextBox1.Text, monthlyTerms) Then
            ' If PaymentMonths reaches zero, reduce monthlyTerms by 1
            If paymentMonths = 0 Then
                monthlyTerms -= 1
            End If

            ' If the updated amount becomes zero, set monthlyTerms to zero
            If updatedAmount = 0 Then
                monthlyTerms = 0
            End If
        Else
            MessageBox.Show("Invalid value for monthly terms. Please enter a valid number.")
            Exit Sub
        End If
        ' Update Amount, Payment, Date, PaymentMonths, and monthly_terms in the Member table
        Try
            SqlConn.Open()
            Dim updateQuery As String = "UPDATE dbo.Member SET [amount] = @amount, [payment] = @payment, [date] = @date, [PaymentMonths] = @paymentMonths, [monthly_terms] = @monthlyTerms WHERE [ID] = @id"
            Using updateCommand As New SqlCommand(updateQuery, SqlConn)
                updateCommand.Parameters.AddWithValue("@id", IDTextBox.Text)
                updateCommand.Parameters.AddWithValue("@amount", updatedAmount)
                updateCommand.Parameters.AddWithValue("@payment", paymentValue) ' Save payment in the payment column
                updateCommand.Parameters.AddWithValue("@date", paymentDate) ' Save the date
                updateCommand.Parameters.AddWithValue("@paymentMonths", paymentMonths) ' Update PaymentMonths
                updateCommand.Parameters.AddWithValue("@monthlyTerms", monthlyTerms) ' Update monthly_terms from TextBox1

                Dim rowsAffected As Integer = updateCommand.ExecuteNonQuery()
                If rowsAffected > 0 Then
                    MessageBox.Show("Amount, payment, date, PaymentMonths, and monthly_terms updated successfully!")

                    Dim insertQuery As String = "INSERT INTO dbo.PaymentHistory ([Payment], [date], [PaymentHistoryID]) VALUES (@payment, @date, @paymentHistoryID)"
                    Using insertCommand As New SqlCommand(insertQuery, SqlConn)
                        insertCommand.Parameters.AddWithValue("@payment", paymentValue) ' Insert the payment
                        insertCommand.Parameters.AddWithValue("@date", paymentDate) ' Date of the payment
                        insertCommand.Parameters.AddWithValue("@paymentHistoryID", IDTextBox.Text) ' Link to the member's ID

                        Dim insertRows As Integer = insertCommand.ExecuteNonQuery()
                        If insertRows > 0 Then
                            MessageBox.Show("Payment recorded in payment history successfully.")
                        Else
                            MessageBox.Show("Failed to record payment in payment history.")
                        End If
                    End Using

                    LoadData()
                Else
                    MessageBox.Show("Failed to update the amount. Ensure the record with the specified ID exists.")
                End If
            End Using
        Catch ex As SqlException
            MessageBox.Show("A database error occurred: " & ex.Message)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            If SqlConn.State = ConnectionState.Open Then
                SqlConn.Close()
            End If
        End Try

        PaymentTextBox.Clear()
    End Sub







    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Search_TextChanged(sender As Object, e As EventArgs) Handles Search.TextChanged
        Dim searchText As String = Search.Text.Trim()

        ' Ensure that the DataGridView has a DataSource
        If DataGridView1.DataSource IsNot Nothing Then
            Dim dataTable As DataTable = CType(DataGridView1.DataSource, DataTable)

            ' Apply the filter to the DefaultView of the DataTable
            If String.IsNullOrEmpty(searchText) Then
                ' If search text is empty, show all rows
                dataTable.DefaultView.RowFilter = String.Empty
            Else
                ' Try to parse the search text as an integer
                Dim idValue As Integer
                If Integer.TryParse(searchText, idValue) Then
                    ' Create the filter string for the DataTable
                    ' Using '=' for integer comparison
                    Dim filterString As String = String.Format("[ID] = {0}", idValue)
                    dataTable.DefaultView.RowFilter = filterString
                Else

                    dataTable.DefaultView.RowFilter = "1=0"
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim currentAmount As Double
        Dim interestRate As Double
        Dim interestAmount As Double
        Dim newAmount As Double

        ' Validate and convert the amount and interest rate to Double
        If Not Double.TryParse(AmountTextBox.Text, currentAmount) Then
            MessageBox.Show("Invalid amount entered in AmountTextBox. Please enter a valid number.")
            Exit Sub
        End If

        If Not Double.TryParse(InterestTextBox.Text, interestRate) Then
            MessageBox.Show("Invalid interest rate entered in InterestTextBox. Please enter a valid number.")
            Exit Sub
        End If

        ' Calculate the interest amount based on the interest rate
        ' Convert the percentage to a decimal by dividing by 100
        interestAmount = currentAmount * (interestRate / 100)

        ' Calculate the new total amount
        newAmount = currentAmount + interestAmount

        ' Display the message with the interest amount and the new total amount
        MessageBox.Show($"The interest amount being added is {interestAmount.ToString("F2")}. The new total amount would be {newAmount.ToString("F2")}.", "Amount Information")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Validate that ID is not empty
        If String.IsNullOrWhiteSpace(IDTextBox.Text) Then
            MessageBox.Show("ID cannot be empty. Please enter a valid ID.")
            Exit Sub
        End If

        ' Create a new instance of the PaymentHistoryForm
        Dim paymentHistoryForm As New PaymentHistoryForm()

        ' Set the MemberId property to the current ID from the TextBox
        paymentHistoryForm.MemberId = IDTextBox.Text

        ' Show the PaymentHistoryForm
        paymentHistoryForm.ShowDialog()
    End Sub




    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DatePicker_ValueChanged(sender As Object, e As EventArgs) Handles DatePicker.ValueChanged

    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim newText As String = TextBox1.Text

        ' Use a StringBuilder to create a new string with only valid characters
        Dim validText As New System.Text.StringBuilder()

        For Each c As Char In newText
            ' Check if the character is a digit or a negative sign
            If Char.IsDigit(c) OrElse (c = "-"c AndAlso validText.Length = 0) Then
                validText.Append(c)
            End If
        Next

        ' Update the TextBox with the valid characters only
        TextBox1.Text = validText.ToString()

        ' Set the cursor position to the end of the TextBox
        TextBox1.SelectionStart = TextBox1.Text.Length
    End Sub


    Private Sub IDTextBox_TextChanged(sender As Object, e As EventArgs) Handles IDTextBox.TextChanged

    End Sub

    Private Sub MemberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles awit.Click
        Me.Hide()
        member.Show()

    End Sub

    Private Sub ContributionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContributionToolStripMenuItem.Click
        Me.Hide()
        Contribution.Show()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim amount As Decimal ' Change to Decimal to allow float numbers
        ' Try to parse the amount from the AmountTextBox
        If Decimal.TryParse(AmountTextBox.Text, amount) Then
            Dim interest As Integer = CalculateInterest(amount)
            InterestTextBox.Text = interest.ToString() ' Display interest as string

            ' Display a note regarding interest calculation
            MessageBox.Show("Note: 1% interest is added per 1000 currency units, with a maximum limit of 15%.", "Interest Calculation Note", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Please enter a valid amount.")
        End If
    End Sub

    Private Function CalculateInterest(amount As Decimal) As Integer
        Dim interest As Integer = 0

        If amount > 0 Then
            ' Base interest: 1% for each 1000
            interest = Math.Floor(amount / 1000) ' 1% for every 1000

            ' Limit the interest to a maximum of 15%
            If interest > 15 Then
                interest = 15
            End If
        End If

        Return interest
    End Function


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Get the ID from IDTextBox to find the corresponding record
        Dim idText As String = IDTextBox.Text
        Dim dateValue As Date = Date.Now ' Current date
        Dim amountValue As Double
        Dim penaltyAmount As Double = 0 ' Initialize penalty amount

        ' Validate and convert amount to Double
        If Not Double.TryParse(AmountTextBox.Text, amountValue) Then
            MessageBox.Show("Invalid amount entered. Please enter a valid number.")
            Exit Sub
        End If

        ' Fetch the existing record's date for calculation
        Dim connectionString As String = "Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1"
        Dim query As String = "SELECT [date], [amount] FROM dbo.Member WHERE [ID] = @id"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@id", idText)

                Try
                    connection.Open()
                    Dim reader As SqlDataReader = command.ExecuteReader()

                    ' Read the data and store it in local variables
                    Dim existingDate As Date
                    Dim existingAmount As Double

                    If reader.Read() Then
                        existingDate = Convert.ToDateTime(reader("date"))
                        existingAmount = Convert.ToDouble(reader("amount"))
                    Else
                        MessageBox.Show("Record not found.")
                        Return ' Exit if the record is not found
                    End If


                    reader.Close()


                    Dim overdueDays As Integer = (dateValue - existingDate).Days
                    MessageBox.Show("Overdue Days: " & overdueDays.ToString()) ' Debugging line

                    If overdueDays > 30 Then
                        ' Validate the penalty amount before applying it
                        If Not Double.TryParse(Penaltytxt.Text, penaltyAmount) Then
                            MessageBox.Show("Invalid penalty amount. Please enter a valid number.")
                            Exit Sub
                        End If

                        ' Calculate penalty
                        Dim extraDays As Integer = overdueDays - 30
                        penaltyAmount = extraDays * penaltyAmount ' Use penalty from the textbox
                        Penaltytxt.ReadOnly = False
                    Else
                        ' Disable PenaltyTextBox if within the penalty month
                        Penaltytxt.ReadOnly = True
                        Penaltytxt.Clear() ' Clear any previous penalty value
                    End If

                    ' Update the amount with the penalty
                    Dim newAmount As Double = existingAmount + penaltyAmount

                    ' Update the AmountTextBox with the new calculated amount
                    AmountTextBox.Text = newAmount.ToString("F2") ' Format to 2 decimal places

                    ' Optional: Update the record in the database
                    Dim updateQuery As String = "UPDATE dbo.Member SET [date] = @newDate, [amount] = @newAmount WHERE [ID] = @id"
                    Using updateCommand As New SqlCommand(updateQuery, connection)
                        updateCommand.Parameters.AddWithValue("@id", idText)
                        updateCommand.Parameters.AddWithValue("@newDate", dateValue)
                        updateCommand.Parameters.AddWithValue("@newAmount", newAmount)

                        Dim rowsAffected As Integer = updateCommand.ExecuteNonQuery()
                        If rowsAffected > 0 Then
                            MessageBox.Show("Record updated successfully with penalties applied!")
                        Else
                            MessageBox.Show("Failed to update the record.")
                        End If
                    End Using
                Catch ex As SqlException
                    MessageBox.Show("A database error occurred: " & ex.Message)
                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message & vbCrLf & "Stack Trace: " & ex.StackTrace)
                Finally
                    connection.Close()
                End Try
            End Using
        End Using
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub SavingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavingsToolStripMenuItem.Click
        Me.Hide()
        Savings.Show()
    End Sub

    Private Sub HomeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HomeToolStripMenuItem.Click
        Me.Hide()
        Form3.Show()

    End Sub

    Private Sub Penaltytxt_TextChanged(sender As Object, e As EventArgs) Handles Penaltytxt.TextChanged
        Dim newText As String = Penaltytxt.Text

        ' Use a StringBuilder to create a new string with only valid characters
        Dim validText As New System.Text.StringBuilder()

        For Each c As Char In newText
            ' Check if the character is a digit or a negative sign
            If Char.IsDigit(c) OrElse (c = "-"c AndAlso validText.Length = 0) Then
                validText.Append(c)
            End If
        Next

        ' Update the TextBox with the valid characters only
        Penaltytxt.Text = validText.ToString()

        ' Set the cursor position to the end of the TextBox
        Penaltytxt.SelectionStart = TextBox2.Text.Length
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Parse the values from the TextBoxes
        Dim amount As Double
        Dim interest As Double
        Dim monthlyTerms As Double

        If Double.TryParse(AmountTextBox.Text, amount) AndAlso
       Double.TryParse(InterestTextBox.Text, interest) AndAlso
       Double.TryParse(TextBox1.Text, monthlyTerms) Then

            ' Convert interest from percentage to decimal (correct calculation)
            Dim interestDecimal As Double = interest / 100 ' Convert percentage to decimal

            ' Calculate the payment using the formula for monthly payments
            ' Using the formula: P = A * (r(1 + r)^n) / ((1 + r)^n - 1)
            ' where P = Payment, A = Loan amount, r = interest rate per period, n = number of periods
            Dim monthlyRate As Double = interestDecimal / 12 ' Monthly interest rate
            Dim numberOfPayments As Double = monthlyTerms

            Dim payment As Double
            If monthlyRate > 0 Then
                payment = amount * (monthlyRate * Math.Pow(1 + monthlyRate, numberOfPayments)) / (Math.Pow(1 + monthlyRate, numberOfPayments) - 1)
            Else
                payment = amount / numberOfPayments ' In case of zero interest
            End If

            ' Display the result in PaymentTextBox
            PaymentTextBox.Text = payment.ToString("F2")

            ' Save the result to the PaymentMonths column in the database
            Try
                ' Connection string
                Dim connectionString As String = "Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1"

                ' SQL query to update the PaymentMonths column only
                Dim query As String = "UPDATE dbo.member SET PaymentMonths = @Payment WHERE ID = @ID"

                ' Create a connection and command
                Using connection As New SqlClient.SqlConnection(connectionString)
                    Using command As New SqlClient.SqlCommand(query, connection)
                        ' Add parameters to avoid SQL injection
                        command.Parameters.AddWithValue("@Payment", payment)
                        command.Parameters.AddWithValue("@ID", Convert.ToInt32(IDTextBox.Text)) ' Use the correct ID value from the form

                        ' Open the connection
                        connection.Open()

                        ' Execute the query
                        Dim rowsAffected As Integer = command.ExecuteNonQuery()
                        If rowsAffected > 0 Then
                            MessageBox.Show("Payment saved successfully to PaymentMonths.")
                        Else
                            MessageBox.Show("No record was updated. Please check the ID.")
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("An error occurred while saving the payment: " & ex.Message)
            End Try
        Else
            MessageBox.Show("Please enter valid numeric values.")
        End If
        LoadData()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim newText As String = TextBox2.Text

        ' Use a StringBuilder to create a new string with only valid characters
        Dim validText As New System.Text.StringBuilder()

        Dim decimalAdded As Boolean = False ' To track if a decimal point has been added

        For Each c As Char In newText
            ' Check if the character is a digit, a negative sign (at the beginning), or a decimal point (only one allowed)
            If Char.IsDigit(c) Then
                validText.Append(c)
            ElseIf c = "-"c AndAlso validText.Length = 0 Then
                validText.Append(c) ' Allow negative sign, but only at the start
            ElseIf c = "."c AndAlso Not decimalAdded Then
                validText.Append(c) ' Allow only one decimal point
                decimalAdded = True
            End If
        Next

        ' Update the TextBox with the valid characters only
        TextBox2.Text = validText.ToString()

        ' Set the cursor position to the end of the TextBox
        TextBox2.SelectionStart = TextBox2.Text.Length
    End Sub
End Class
