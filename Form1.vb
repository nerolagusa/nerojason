Imports System.Data.SqlClient

Public Class Form1
    Dim SqlConn As SqlConnection
    Dim COMMAND As SqlCommand

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SqlConn = New SqlConnection("Data Source=COMP64\SQLEXPRESS;Initial Catalog=renniel;Persist Security Info=True;User ID=login1;Password=renniel1")
        TextBox2.PasswordChar = "*" ' Set the initial password character
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim username As String = TextBox1.Text
        Dim password As String = TextBox2.Text

        Try
            Using SqlConn
                SqlConn.Open()

                Dim query As String = "SELECT COUNT(*) FROM dbo.login WHERE USERNAME = @username AND PASSWORD = @password"
                Using command As New SqlCommand(query, SqlConn)
                    command.Parameters.AddWithValue("@username", username)
                    command.Parameters.AddWithValue("@password", password)

                    Dim result As Integer = Convert.ToInt32(command.ExecuteScalar())

                    If result > 0 Then
                        MessageBox.Show("Login successful")
                        Me.Hide()
                        Form3.Show()
                    Else
                        MessageBox.Show("Invalid username or password")
                    End If
                End Using
            End Using
        Catch ex As SqlException
            MessageBox.Show("A database error occurred: " & ex.Message)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ' Toggle the PasswordChar based on the CheckBox state
        If CheckBox1.Checked Then
            TextBox2.PasswordChar = "" ' Show the password
        Else
            TextBox2.PasswordChar = "*" ' Hide the password
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Hide()
        Form2.Show()

    End Sub
End Class
