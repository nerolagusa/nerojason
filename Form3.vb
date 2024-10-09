Public Class Form3
    Private Sub ContributionToolStripMenuItem_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub LoanToolStripMenuItem_Click(sender As Object, e As EventArgs) 
        Me.Hide()
        Loan.Show()


    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        member.Show()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Contribution.Show()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Loan.Show()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
        Savings.Show()

    End Sub
End Class