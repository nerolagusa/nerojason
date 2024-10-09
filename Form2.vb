Imports System.Data.SqlClient

Public Class Form2
    Dim SqlConn As SqlConnection
    Dim COMMAND As SqlCommand

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize the connection string
        SqlConn = New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Get the values from TextBoxes
        Dim username As String = TextBox1.Text
        Dim password As String = TextBox2.Text

        Try
            ' Open the SQL connection
            Using SqlConn
                SqlConn.Open()

                ' SQL Insert Query to add a new record
                Dim insertQuery As String = "INSERT INTO dbo.login (USERNAME, PASSWORD) VALUES (@username, @password)"
                Using command As New SqlCommand(insertQuery, SqlConn)
                    ' Add parameters to prevent SQL injection
                    command.Parameters.AddWithValue("@username", username)
                    command.Parameters.AddWithValue("@password", password)

                    ' Execute the query
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()

                    ' Check if the insertion was successful
                    If rowsAffected > 0 Then
                        MessageBox.Show("Record saved successfully!")
                    Else
                        MessageBox.Show("Failed to save the record.")
                    End If
                End Using
            End Using
        Catch ex As SqlException
            ' Handle SQL exception
            MessageBox.Show("A database error occurred: " & ex.Message)
        Catch ex As Exception
            ' Handle general exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.Hide()
        Form1.Show()
    End Sub
End Class