Imports System.Data.SqlClient
Imports QRCoder
Public Class Contribution
    Dim SqlConn As SqlConnection
    Dim COMMAND As SqlCommand
    Dim dataAdapter As SqlDataAdapter
    Dim dataTable As DataTable

    Private Sub Contribution_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize the connection to the SQL Server database
        SqlConn = New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
        Dim idValue As Integer
        If Integer.TryParse(TextBox4.Text, idValue) Then
            RefreshData(idValue)
        End If

    End Sub









    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ' Handle text change for TextBox1 (Contribution amount)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        ' Handle text change for another TextBox if needed
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)
        ' Handle text change for another TextBox if needed
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Open the SQL connection if it's closed
        If SqlConn.State = ConnectionState.Closed Then
            SqlConn.Open()
        End If

        Try
            ' Retrieve the Contribution value from TextBox3 and the ID from TextBox4
            Dim contributionValue As Decimal
            If Not Decimal.TryParse(TextBox5.Text, contributionValue) Then
                MessageBox.Show("Invalid contribution amount. Please enter a valid number.")
                Exit Sub
            End If

            Dim idValue As Integer
            If Not Integer.TryParse(TextBox4.Text, idValue) Then
                MessageBox.Show("Invalid ID. Please enter a valid number.")
                Exit Sub
            End If

            ' Check the current total amount before proceeding
            Dim currentTotal As Decimal = 0
            Dim totalQuery As String = "SELECT Total FROM dbo.member WHERE ID = @ID"

            Using totalCommand As New SqlCommand(totalQuery, SqlConn)
                totalCommand.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and retrieve the current total
                Dim result = totalCommand.ExecuteScalar()
                If result IsNot Nothing Then
                    currentTotal = Convert.ToDecimal(result)
                End If
            End Using

            ' Check if contribution exceeds the current total
            If contributionValue > currentTotal Then
                MessageBox.Show("Insufficient amount. Contribution cannot exceed the total.")
                Exit Sub
            End If

            ' SQL Update Query to update the Contribution and subtract it from the Total in the dbo.member table
            Dim updateQuery As String = "UPDATE dbo.member " &
                                "SET Contribution = @Contribution, " &
                                "Total = Total - @Contribution " &
                                "WHERE ID = @ID"

            ' Use SqlCommand to execute the query
            Using COMMAND As New SqlCommand(updateQuery, SqlConn)
                ' Add the contribution and ID parameters to the SQL command
                COMMAND.Parameters.AddWithValue("@Contribution", contributionValue)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query
                Dim rowsAffected As Integer = COMMAND.ExecuteNonQuery()

                ' Check if the update was successful
                If rowsAffected > 0 Then
                    MessageBox.Show("Contribution updated and deducted from Total successfully!")
                Else
                    MessageBox.Show("Failed to update contribution. Ensure the ID is correct.")
                End If
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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Open the SQL connection if it's closed
        If SqlConn.State = ConnectionState.Closed Then
            SqlConn.Open()
        End If

        Try
            ' Retrieve the Contribution value from TextBox3 and the ID from TextBox4
            Dim contributionValue As Decimal
            If Not Decimal.TryParse(TextBox7.Text, contributionValue) Then
                MessageBox.Show("Invalid contribution amount. Please enter a valid number.")
                Exit Sub
            End If

            Dim idValue As Integer
            If Not Integer.TryParse(TextBox4.Text, idValue) Then
                MessageBox.Show("Invalid ID. Please enter a valid number.")
                Exit Sub
            End If

            ' Check the current total amount before proceeding
            Dim currentTotal As Decimal = 0
            Dim totalQuery As String = "SELECT Total FROM dbo.member WHERE ID = @ID"

            Using totalCommand As New SqlCommand(totalQuery, SqlConn)
                totalCommand.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and retrieve the current total
                Dim result = totalCommand.ExecuteScalar()
                If result IsNot Nothing Then
                    currentTotal = Convert.ToDecimal(result)
                End If
            End Using

            ' Check if contribution exceeds the current total
            If contributionValue > currentTotal Then
                MessageBox.Show("Insufficient amount. Contribution cannot exceed the total.")
                Exit Sub
            End If

            ' SQL Update Query to update the Contribution and subtract it from the Total in the dbo.member table
            Dim updateQuery As String = "UPDATE dbo.member " &
                                "SET Contribution = @Contribution, " &
                                "Total = Total - @Contribution " &
                                "WHERE ID = @ID"

            ' Use SqlCommand to execute the query
            Using COMMAND As New SqlCommand(updateQuery, SqlConn)
                ' Add the contribution and ID parameters to the SQL command
                COMMAND.Parameters.AddWithValue("@Contribution", contributionValue)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query
                Dim rowsAffected As Integer = COMMAND.ExecuteNonQuery()

                ' Check if the update was successful
                If rowsAffected > 0 Then
                    MessageBox.Show("Contribution updated and deducted from Total successfully!")
                Else
                    MessageBox.Show("Failed to update contribution. Ensure the ID is correct.")
                End If
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

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Open the SQL connection if it's closed
        If SqlConn.State = ConnectionState.Closed Then
            SqlConn.Open()
        End If

        Try
            ' Retrieve the Contribution value from TextBox3 and the ID from TextBox4
            Dim contributionValue As Decimal
            If Not Decimal.TryParse(TextBox8.Text, contributionValue) Then
                MessageBox.Show("Invalid contribution amount. Please enter a valid number.")
                Exit Sub
            End If

            Dim idValue As Integer
            If Not Integer.TryParse(TextBox4.Text, idValue) Then
                MessageBox.Show("Invalid ID. Please enter a valid number.")
                Exit Sub
            End If

            ' Check the current total amount before proceeding
            Dim currentTotal As Decimal = 0
            Dim totalQuery As String = "SELECT Total FROM dbo.member WHERE ID = @ID"

            Using totalCommand As New SqlCommand(totalQuery, SqlConn)
                totalCommand.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and retrieve the current total
                Dim result = totalCommand.ExecuteScalar()
                If result IsNot Nothing Then
                    currentTotal = Convert.ToDecimal(result)
                End If
            End Using

            ' Check if contribution exceeds the current total
            If contributionValue > currentTotal Then
                MessageBox.Show("Insufficient amount. Contribution cannot exceed the total.")
                Exit Sub
            End If

            ' SQL Update Query to update the Contribution and subtract it from the Total in the dbo.member table
            Dim updateQuery As String = "UPDATE dbo.member " &
                                "SET Contribution = @Contribution, " &
                                "Total = Total - @Contribution " &
                                "WHERE ID = @ID"

            ' Use SqlCommand to execute the query
            Using COMMAND As New SqlCommand(updateQuery, SqlConn)
                ' Add the contribution and ID parameters to the SQL command
                COMMAND.Parameters.AddWithValue("@Contribution", contributionValue)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query
                Dim rowsAffected As Integer = COMMAND.ExecuteNonQuery()

                ' Check if the update was successful
                If rowsAffected > 0 Then
                    MessageBox.Show("Contribution updated and deducted from Total successfully!")
                Else
                    MessageBox.Show("Failed to update contribution. Ensure the ID is correct.")
                End If
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Open the SQL connection if it's closed
        If SqlConn.State = ConnectionState.Closed Then
            SqlConn.Open()
        End If

        Try
            ' Retrieve the Contribution value from TextBox3 and the ID from TextBox4
            Dim contributionValue As Decimal
            If Not Decimal.TryParse(TextBox6.Text, contributionValue) Then
                MessageBox.Show("Invalid contribution amount. Please enter a valid number.")
                Exit Sub
            End If

            Dim idValue As Integer
            If Not Integer.TryParse(TextBox4.Text, idValue) Then
                MessageBox.Show("Invalid ID. Please enter a valid number.")
                Exit Sub
            End If

            ' Check the current total amount before proceeding
            Dim currentTotal As Decimal = 0
            Dim totalQuery As String = "SELECT Total FROM dbo.member WHERE ID = @ID"

            Using totalCommand As New SqlCommand(totalQuery, SqlConn)
                totalCommand.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and retrieve the current total
                Dim result = totalCommand.ExecuteScalar()
                If result IsNot Nothing Then
                    currentTotal = Convert.ToDecimal(result)
                End If
            End Using

            ' Check if contribution exceeds the current total
            If contributionValue > currentTotal Then
                MessageBox.Show("Insufficient amount. Contribution cannot exceed the total.")
                Exit Sub
            End If

            ' SQL Update Query to update the Contribution and subtract it from the Total in the dbo.member table
            Dim updateQuery As String = "UPDATE dbo.member " &
                                "SET Contribution = @Contribution, " &
                                "Total = Total - @Contribution " &
                                "WHERE ID = @ID"

            ' Use SqlCommand to execute the query
            Using COMMAND As New SqlCommand(updateQuery, SqlConn)
                ' Add the contribution and ID parameters to the SQL command
                COMMAND.Parameters.AddWithValue("@Contribution", contributionValue)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query
                Dim rowsAffected As Integer = COMMAND.ExecuteNonQuery()

                ' Check if the update was successful
                If rowsAffected > 0 Then
                    MessageBox.Show("Contribution updated and deducted from Total successfully!")
                Else
                    MessageBox.Show("Failed to update contribution. Ensure the ID is correct.")
                End If
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

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Open the SQL connection if it's closed
        If SqlConn.State = ConnectionState.Closed Then
            SqlConn.Open()
        End If

        Try
            ' Retrieve the ID from TextBox4
            Dim idValue As Integer
            If Not Integer.TryParse(TextBox4.Text, idValue) Then
                MessageBox.Show("Invalid ID. Please enter a valid number.")
                Exit Sub
            End If

            ' SQL Query to fetch data for the specified ID
            Dim query As String = "SELECT Name, Total, Contribution, Picture FROM dbo.member WHERE ID = @ID"

            Using COMMAND As New SqlCommand(query, SqlConn)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = COMMAND.ExecuteReader()
                    If reader.Read() Then
                        ' Populate the TextBoxes and Label with retrieved values
                        TextBox1.Text = reader("Name").ToString()                     ' Name
                        TextBox2.Text = reader("Total").ToString()                    ' Total

                        ' Display the total in Label4
                        Label4.Text = "Total: " & reader("Total").ToString()         ' Set total in Label4

                        ' Retrieve the image from the database
                        If Not IsDBNull(reader("Picture")) Then
                            Dim imgData As Byte() = CType(reader("Picture"), Byte())
                            Using ms As New System.IO.MemoryStream(imgData)
                                PictureBox1.Image = Image.FromStream(ms)
                            End Using
                            ' Set the PictureBox SizeMode to StretchImage
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                        Else
                            PictureBox1.Image = Nothing  ' Clear the PictureBox if no image
                        End If

                        ' Clear PictureBox2 since QR code retrieval is removed

                    Else
                        MessageBox.Show("No record found for the provided ID.")
                    End If
                End Using
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

    Private Sub TextBox4_TextChanged_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub LoanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoanToolStripMenuItem.Click
        Me.Hide()
        Loan.Show()
    End Sub

    Private Sub SavingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavingsToolStripMenuItem.Click
        Me.Hide()
        Savings.Show()

    End Sub

    Private Sub HomeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HomeToolStripMenuItem.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub RefreshData(idValue As Integer)
        ' Open the SQL connection if it's closed
        If SqlConn.State = ConnectionState.Closed Then
            SqlConn.Open()
        End If

        Try
            ' SQL Query to fetch updated data for the specified ID
            Dim query As String = "SELECT Name, Total, Contribution, Picture FROM dbo.member WHERE ID = @ID"

            Using COMMAND As New SqlCommand(query, SqlConn)
                COMMAND.Parameters.AddWithValue("@ID", idValue)

                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = COMMAND.ExecuteReader()
                    If reader.Read() Then
                        ' Update the TextBoxes with the retrieved values
                        TextBox1.Text = reader("Name").ToString()              ' Name
                        TextBox2.Text = reader("Total").ToString()             ' Total

                        ' Display the updated total in Label4
                        Label4.Text = "Total: " & reader("Total").ToString()

                        ' Retrieve the image from the database and update the PictureBox
                        If Not IsDBNull(reader("Picture")) Then
                            Dim imgData As Byte() = CType(reader("Picture"), Byte())
                            Using ms As New System.IO.MemoryStream(imgData)
                                PictureBox1.Image = Image.FromStream(ms)
                            End Using
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                        Else
                            PictureBox1.Image = Nothing ' Clear PictureBox if no image exists
                        End If
                    Else
                        MessageBox.Show("No record found for the provided ID.")
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Close the connection if it's open
            If SqlConn.State = ConnectionState.Open Then
                SqlConn.Close()
            End If
        End Try
    End Sub

    Private Sub Member_Click(sender As Object, e As EventArgs)

    End Sub
End Class
