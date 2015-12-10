Imports System.Data.SqlClient

Public Class LoginForm1


    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If (ComboBox1.GetItemText(ComboBox1.SelectedItem).Length > 0) Then

            Hotel_Book.Enabled = True
            Hotel_Book.Label84.Text = getEID()
            Hotel_Book.TextBox55.Text = Hotel_Book.Label84.Text.ToString
            Hotel_Book.TextBox28.Text = Hotel_Book.Label84.Text.ToString
            Me.Close()
        End If
    End Sub

    Public Function getEID() As String
        getEID = (ComboBox1.GetItemText(ComboBox1.SelectedItem))
    End Function

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
        Hotel_Book.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterScreen
        ControlBox = False
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadEmployees(con)
    End Sub

    Private Sub LoadEmployees(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Distinct E_ID From Employee", connection)
            adapter.Fill(ds)
            ComboBox1.DataSource = ds.Tables(0)
            ComboBox1.DisplayMember = "E_ID"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadEmployees")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select Distinct E_ID From Employee"
                command.ExecuteNonQuery()
                transaction.Commit()

            Catch ex As Exception
                Console.WriteLine("Commit Exception Type: {0}", ex.GetType())
                Console.WriteLine("  Message: {0}", ex.Message)

                Try
                    transaction.Rollback()

                Catch ex2 As Exception
                    Console.WriteLine("Rollback Exception Type: {0}", ex2.GetType())
                    Console.WriteLine("  Message: {0}", ex2.Message)
                End Try
            End Try
        End Using
    End Sub
End Class
