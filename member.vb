Imports System.Data.SqlClient
Imports System.IO
Imports QRCoder
Imports System.Drawing
Public Class member
    Dim SqlConn As SqlConnection
    Dim COMMAND As SqlCommand

    Private Sub member_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize the connection string
        SqlConn = New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
        LoadData()
    End Sub

    Private Sub LoadData()
        ' Create a connection string (adjust it according to your database)
        Dim SqlConn As New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")

        Try
            ' Open the connection
            SqlConn.Open()

            ' SQL Query to select all members from the database
            Dim query As String = "SELECT ID, Name, [date], Picture, qrcode FROM dbo.member" ' Make sure to include qrcode here

            ' Create a DataAdapter to fetch the data
            Dim adapter As New SqlDataAdapter(query, SqlConn)

            ' Create a DataTable to store the data
            Dim dataTable As New DataTable()

            ' Fill the DataTable with the data from the adapter
            adapter.Fill(dataTable)

            ' Remove rows where any of the key columns (ID, Name, [date], Picture) are empty or null
            For Each row As DataRow In dataTable.Rows.Cast(Of DataRow)().ToList()
                If row.ItemArray.Any(Function(value) IsDBNull(value) OrElse (TypeOf value Is String AndAlso String.IsNullOrWhiteSpace(value.ToString()))) Then
                    dataTable.Rows.Remove(row)
                End If
            Next

            ' Bind the filtered DataTable to the DataGridView
            DataGridView1.DataSource = dataTable

            ' Configure the DataGridView size
            DataGridView1.Size = New Size(700, 180)

            ' Ensure columns fill the DataGridView width
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' Adjust rows to fit content but preserve row height
            DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            ' Remove unwanted columns
            For Each column As DataGridViewColumn In DataGridView1.Columns.Cast(Of DataGridViewColumn)().ToList()
                If Not {"ID", "Name", "date", "Picture", "qrcode"}.Contains(column.HeaderText) Then ' Include qrcode if you want to access it later
                    DataGridView1.Columns.Remove(column)
                End If
            Next

            ' Configure the DataGridView to display images correctly
            For Each column As DataGridViewColumn In DataGridView1.Columns
                If column.HeaderText = "Picture" Then
                    Dim imageColumn As DataGridViewImageColumn = CType(column, DataGridViewImageColumn)
                    imageColumn.ImageLayout = DataGridViewImageCellLayout.Stretch ' Stretch the image to fit the cell
                    imageColumn.Width = 100 ' Adjust as needed
                End If
            Next

            ' Hide the extra row (new row)
            DataGridView1.AllowUserToAddRows = False

        Catch ex As Exception
            MessageBox.Show("An error occurred while loading data: " & ex.Message)
        Finally
            ' Close the connection
            SqlConn.Close()
        End Try
        Me.qrcodePic.Size = New Size(276, 196)

    End Sub


    Private Function DoesIdExist(id As Integer) As Boolean
        Dim exists As Boolean = False
        Try
            ' Open the connection
            SqlConn.Open()

            ' SQL Query to check if the ID exists
            Dim query As String = "SELECT COUNT(*) FROM dbo.member WHERE ID = @id"
            Using command As New SqlCommand(query, SqlConn)
                command.Parameters.AddWithValue("@id", id)
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                exists = count > 0
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while checking the ID: " & ex.Message)
        Finally
            ' Close the connection
            If SqlConn.State = ConnectionState.Open Then
                SqlConn.Close()
            End If
        End Try
        Return exists
    End Function

    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click
        ' Retrieve the member ID and other details
        Dim memberId As Integer = Convert.ToInt32(TextBox4.Text) ' Assuming TextBox4 contains the ID
        Dim name As String = TextBox2.Text
        Dim memberDate As Date = DateTimePicker1.Value

        ' Generate QR code for the member ID
        Dim qrCodeImage As Image = GenerateQRCode(memberId.ToString())
        Dim qrCodeData As Byte() = Nothing

        ' Check if the QR code image was generated
        If qrCodeImage IsNot Nothing Then
            Try
                ' Convert the QR code image to a byte array safely
                Using ms As New MemoryStream()
                    qrCodeImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    qrCodeData = ms.ToArray()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error converting QR code image: " & ex.Message)
                Exit Sub
            End Try
        End If

        ' Check if an image is loaded in PictureBox1
        Dim imageData As Byte() = Nothing
        If PictureBox1.Image IsNot Nothing Then
            Try
                ' Convert the image in PictureBox1 to a byte array safely
                Using ms As New MemoryStream()
                    PictureBox1.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    imageData = ms.ToArray()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error converting image: " & ex.Message)
                Exit Sub
            End Try
        End If

        ' SQL Insert Query
        Dim query As String = "INSERT INTO dbo.member (ID, Name, [date], Picture, qrcode) VALUES (@id, @name, @date, @picture, @qrcode)"

        ' Create SQL Connection
        Using SqlConn As New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
            Try
                ' Open the connection
                SqlConn.Open()

                ' Create SQL Command
                Using command As New SqlCommand(query, SqlConn)
                    ' Add parameters to prevent SQL injection
                    command.Parameters.AddWithValue("@id", memberId)
                    command.Parameters.AddWithValue("@name", name)
                    command.Parameters.AddWithValue("@date", memberDate)

                    ' Add the image parameter; if no image, set to DBNull
                    If imageData IsNot Nothing Then
                        command.Parameters.AddWithValue("@picture", imageData)
                    Else
                        command.Parameters.AddWithValue("@picture", DBNull.Value)
                    End If

                    ' Add the QR code image parameter; if no QR code image, set to DBNull
                    If qrCodeData IsNot Nothing Then
                        command.Parameters.AddWithValue("@qrcode", qrCodeData)
                    Else
                        command.Parameters.AddWithValue("@qrcode", DBNull.Value)
                    End If

                    ' Execute the query
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()

                    ' Check if the insert was successful
                    If rowsAffected > 0 Then
                        MessageBox.Show("Member record inserted successfully!")
                    Else
                        MessageBox.Show("Failed to insert the member record.")
                    End If
                End Using

            Catch ex As SqlException
                ' Handle SQL exception
                MessageBox.Show("A database error occurred: " & ex.Message)
            Catch ex As Exception
                ' Handle general exception
                MessageBox.Show("An error occurred: " & ex.Message)
            End Try
        End Using
    End Sub

    ' Function to generate QR code from text
    Private Function GenerateQRCode(code As String) As Image
        ' Create a new instance of the QR code generator
        Dim qrGenerator As New QRCodeGenerator()
        Dim qrCodeData As QRCodeData = qrGenerator.CreateQrCode(code, QRCodeGenerator.ECCLevel.Q)
        Dim qrCode As New QRCode(qrCodeData)

        ' Generate the QR code as a Bitmap
        Return qrCode.GetGraphic(20)
    End Function


    Private Function ResizeImage(ByVal image As Image, ByVal width As Integer, ByVal height As Integer) As Image
        Dim resizedImage As New Bitmap(width, height)
        Using graphics As Graphics = Graphics.FromImage(resizedImage)
            graphics.DrawImage(image, 0, 0, width, height)
        End Using
        Return resizedImage
    End Function



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Get values from TextBoxes, DateTimePicker, and PictureBox
        Dim memberId As Integer
        If Not Integer.TryParse(TextBox4.Text, memberId) Then
            MessageBox.Show("Please enter a valid ID.")
            Return
        End If

        Dim name As String = TextBox2.Text
        Dim memberDate As Date = DateTimePicker1.Value

        ' Check if an image is loaded in PictureBox1
        Dim imageData As Byte() = Nothing
        If PictureBox1.Image IsNot Nothing Then
            Try
                Using ms As New MemoryStream()
                    PictureBox1.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    imageData = ms.ToArray()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error converting image: " & ex.Message)
                Return
            End Try
        End If

        ' Check if a QR code image is loaded in qrcodePic
        Dim qrCodeData As Byte() = Nothing
        If qrcodePic.Image IsNot Nothing Then
            Try
                Using ms As New MemoryStream()
                    qrcodePic.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    qrCodeData = ms.ToArray()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error converting QR code image: " & ex.Message)
                Return
            End Try
        End If

        ' SQL Update Query
        Dim query As String = "UPDATE dbo.member SET Name = @name, [date] = @date, Picture = @picture, qrcode = @qrcode WHERE ID = @id"

        ' Create SQL Connection
        Using SqlConn As New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
            Try
                SqlConn.Open()

                ' Create SQL Command
                Using command As New SqlCommand(query, SqlConn)
                    ' Add parameters to prevent SQL injection
                    command.Parameters.AddWithValue("@id", memberId)
                    command.Parameters.AddWithValue("@name", name)
                    command.Parameters.AddWithValue("@date", memberDate)

                    ' Add the image parameter; if no image, set to DBNull
                    If imageData IsNot Nothing Then
                        command.Parameters.AddWithValue("@picture", imageData)
                    Else
                        command.Parameters.AddWithValue("@picture", DBNull.Value)
                    End If

                    ' Add the QR code image parameter; if no QR code image, set to DBNull
                    If qrCodeData IsNot Nothing Then
                        command.Parameters.AddWithValue("@qrcode", qrCodeData)
                    Else
                        command.Parameters.AddWithValue("@qrcode", DBNull.Value)
                    End If

                    ' Execute the query
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()

                    ' Check if the update was successful
                    If rowsAffected > 0 Then
                        MessageBox.Show("Member record updated successfully!")
                    Else
                        MessageBox.Show("Failed to update the member record. Please ensure the ID exists.")
                    End If
                End Using

            Catch ex As SqlException
                ' Handle SQL exception
                MessageBox.Show("A database error occurred: " & ex.Message)
            Catch ex As Exception
                ' Handle general exception
                MessageBox.Show("An error occurred: " & ex.Message)
            End Try
        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Get the ID from TextBox4
        Dim idValue As Integer
        If Integer.TryParse(TextBox4.Text, idValue) Then
            ' Check if ID is greater than 0
            If idValue > 0 Then
                ' Show confirmation dialog
                Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete the record with ID " & idValue.ToString() & "?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                ' If user confirms deletion
                If result = DialogResult.Yes Then
                    Try
                        ' Open the connection
                        SqlConn.Open()

                        ' Begin a transaction
                        Using transaction As SqlTransaction = SqlConn.BeginTransaction()
                            ' SQL query to delete the record from the member table
                            Dim deleteMemberQuery As String = "DELETE FROM dbo.member WHERE ID = @id"
                            Using memberCommand As New SqlCommand(deleteMemberQuery, SqlConn, transaction)
                                ' Add the parameter to prevent SQL injection
                                memberCommand.Parameters.AddWithValue("@id", idValue)

                                ' Execute the query for deleting member
                                Dim rowsAffected As Integer = memberCommand.ExecuteNonQuery()

                                ' Check if member deletion was successful
                                If rowsAffected > 0 Then
                                    ' SQL query to delete the corresponding payment history
                                    Dim deletePaymentHistoryQuery As String = "DELETE FROM dbo.PaymentHistory WHERE PaymentHistoryID = @id"
                                    Using paymentHistoryCommand As New SqlCommand(deletePaymentHistoryQuery, SqlConn, transaction)
                                        ' Add the parameter for PaymentHistoryID
                                        paymentHistoryCommand.Parameters.AddWithValue("@id", idValue)

                                        ' Execute the query for deleting payment history
                                        Dim paymentHistoryRowsAffected As Integer = paymentHistoryCommand.ExecuteNonQuery()

                                        ' Commit the transaction if both deletions were successful
                                        If paymentHistoryRowsAffected >= 0 Then
                                            transaction.Commit()
                                            MessageBox.Show("Record and associated payment history deleted successfully!")
                                            LoadData() ' Refresh the DataGridView after deletion
                                        Else
                                            MessageBox.Show("No corresponding payment history found for the given ID.")
                                        End If
                                    End Using
                                Else
                                    MessageBox.Show("No record found with the given ID in the member table.")
                                End If
                            End Using
                        End Using
                    Catch ex As SqlException
                        MessageBox.Show("Database error: " & ex.Message)
                    Catch ex As Exception
                        MessageBox.Show("An error occurred: " & ex.Message)
                    Finally
                        ' Close the connection
                        If SqlConn.State = ConnectionState.Open Then
                            SqlConn.Close()
                        End If
                    End Try
                End If
            Else
                MessageBox.Show("Please enter a valid ID.")
            End If
        Else
            MessageBox.Show("Invalid ID format.")
        End If
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Reload data in DataGridView
        LoadData()

        ' Clear or reset controls to default values
        TextBox4.Clear()    ' Clear ID TextBox
        TextBox2.Clear()    ' Clear Name TextBox
        DateTimePicker1.Value = DateTime.Now ' Reset DateTimePicker to current date

        ' Clear PictureBox1
        PictureBox1.Image = Nothing

        ' Clear qrcodePic
        qrcodePic.Image = Nothing

        ' Clear TextBox3 for search
        TextBox3.Clear()
        TextBox4.Enabled = True
        Button5.Enabled = True
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        ' Ensure the click is on a valid row and not the header
        If e.RowIndex >= 0 Then
            ' Get the selected row
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Populate the form controls with the selected row data
            TextBox4.Text = row.Cells("ID").Value.ToString() ' ID
            TextBox2.Text = row.Cells("Name").Value.ToString() ' Name
            DateTimePicker1.Value = Convert.ToDateTime(row.Cells("date").Value) ' Date

            ' Check if the picture is not null before displaying it
            If Not IsDBNull(row.Cells("Picture").Value) Then
                Dim imageData As Byte() = CType(row.Cells("Picture").Value, Byte())
                Using ms As New MemoryStream(imageData)
                    PictureBox1.Image = Image.FromStream(ms)
                End Using
            Else
                PictureBox1.Image = Nothing ' Set to no image if picture is null
            End If

            ' Set the PictureBox size to 145x102 and apply StretchImage to fit the size
            PictureBox1.Size = New Size(145, 102)
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

            ' Check if the QR code image is not null before displaying it
            If Not IsDBNull(row.Cells("qrcode").Value) Then
                Dim qrCodeData As Byte() = CType(row.Cells("qrcode").Value, Byte())
                Using ms As New MemoryStream(qrCodeData)
                    qrcodePic.Image = Image.FromStream(ms)
                End Using
            Else
                qrcodePic.Image = Nothing ' Set to no image if QR code is null
            End If

            ' Set the QR code PictureBox size to 276x196 and apply StretchImage to fit the size
            qrcodePic.Size = New Size(276, 196)
            qrcodePic.SizeMode = PictureBoxSizeMode.StretchImage

            ' Disable the ID TextBox
            TextBox4.Enabled = False
            Button5.Enabled = False
        End If
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim searchText As String = TextBox3.Text.Trim()

        ' Ensure that the DataGridView has a DataSource
        If DataGridView1.DataSource IsNot Nothing Then
            Dim dataTable As DataTable = CType(DataGridView1.DataSource, DataTable)

            ' Apply the filter to the DefaultView of the DataTable
            If String.IsNullOrEmpty(searchText) Then
                ' If search text is empty, show all rows
                dataTable.DefaultView.RowFilter = String.Empty
            Else
                ' Create the filter string for the DataTable to search in the ID column
                Dim filterString As String = String.Format("[ID] = {0}", searchText.Replace("'", "''"))
                dataTable.DefaultView.RowFilter = filterString
            End If
        End If
    End Sub



    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
        ' Allow only digits and the character "-"
        Dim input As String = TextBox4.Text
        Dim validInput As New System.Text.StringBuilder()

        For Each ch As Char In input
            If Char.IsDigit(ch) OrElse ch = "-" Then
                validInput.Append(ch)
            End If
        Next

        ' Update the TextBox with valid input
        TextBox4.Text = validInput.ToString()

        ' Set the cursor position to the end of the TextBox
        TextBox4.SelectionStart = TextBox4.Text.Length
    End Sub



    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        ' Check if the input is valid
        Dim regex As New System.Text.RegularExpressions.Regex("[^\d-]")
        If regex.IsMatch(TextBox4.Text) Then
            ' Remove the invalid character
            TextBox4.Text = regex.Replace(TextBox4.Text, "")
            ' Move the cursor to the end of the text
            TextBox4.SelectionStart = TextBox4.Text.Length
        End If
    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        ' Check if the input contains any digits
        Dim regex As New System.Text.RegularExpressions.Regex("\d")
        If regex.IsMatch(TextBox2.Text) Then
            ' Remove the invalid character(s)
            TextBox2.Text = regex.Replace(TextBox2.Text, "")
            ' Move the cursor to the end of the text
            TextBox2.SelectionStart = TextBox2.Text.Length
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Create a new instance of OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()

        ' Set the properties of the OpenFileDialog
        openFileDialog.Title = "Select an Image"
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif|All Files|*.*"

        ' Show the dialog and check if the user selected a file
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Load the selected image into PictureBox1
            Try
                ' Load the image into PictureBox1
                PictureBox1.Image = Image.FromFile(openFileDialog.FileName)
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage ' Optional: Set the image to fit the PictureBox
            Catch ex As Exception
                MessageBox.Show("Error loading image: " & ex.Message)
            End Try
        End If

    End Sub

    Private Sub LoanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoanToolStripMenuItem.Click
        Me.Hide()
        Loan.Show()

    End Sub

    Private Sub ContributionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContributionToolStripMenuItem.Click
        Me.Hide()
        Contribution.Show()
    End Sub

    Private Sub SavingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavingsToolStripMenuItem.Click
        Me.Hide()
        Savings.Show()

    End Sub

    Private Sub HomeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HomeToolStripMenuItem.Click
        Me.Hide()
        Form3.Show()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Generate a random code (not displayed in QR code)
        Dim random As New Random()
        Dim randomCode As String = random.Next(100000, 999999).ToString()

        ' Create a new instance of the QR code generator
        Dim qrGenerator As New QRCodeGenerator()
        Dim qrCodeData As QRCodeData = qrGenerator.CreateQrCode(" ", QRCodeGenerator.ECCLevel.Q) ' Empty value for QR code
        Dim qrCode As New QRCode(qrCodeData)

        ' Generate the QR code as a Bitmap
        Dim qrCodeImage As Bitmap = qrCode.GetGraphic(20)
        Me.qrcodePic.Size = New Size(276, 196)
        Me.qrcodePic.SizeMode = PictureBoxSizeMode.StretchImage

        ' Display the QR code in the qrcode PictureBox
        Me.qrcodePic.Image = qrCodeImage

        ' Optionally store the random code in a variable if needed later
        ' You can also consider storing it if you need a reference
        ' savedRandomCode = randomCode
    End Sub

    Private Sub qrcode_Click(sender As Object, e As EventArgs) Handles qrcodePic.Click

        qrcodePic.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub
    Private Function DecodeQRCode(qrCodeImage As Image) As String
        Dim barcodeReader As New ZXing.BarcodeReader()
        Dim bitmap As Bitmap = New Bitmap(qrCodeImage)

        Dim result As ZXing.Result = barcodeReader.Decode(bitmap)

        If result IsNot Nothing Then
            Return result.Text ' Return the decoded text
        Else
            Return String.Empty ' Return empty string if decoding fails
        End If
    End Function
End Class
