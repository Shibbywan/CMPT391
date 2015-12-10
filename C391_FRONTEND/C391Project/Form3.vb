Imports System.Data.SqlClient

Public Class customerListForm

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Me.StartPosition = FormStartPosition.CenterScreen
        LoadCustomers(con)
    End Sub

    Private Sub LoadCustomers(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("SELECT C_ID as 'Customer ID', Name FROM Customer, Person WHERE Customer.C_ID = Person.P_ID", connection)
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("Load Customers")
            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "SELECT C_ID as 'Customer ID', Name FROM Customer, Person WHERE Customer.C_ID = Person.P_ID"
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Dim i = DataGridView1.CurrentRow.Index
        Dim cid = DataGridView1.Item(1, i).Value
        Dim name = DataGridView1.Item(0, i).Value
        TextBox2.Text = name
        TextBox3.Text = cid
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (TextBox2.Text.Length > 0) Then
            Hotel_Book.cidField.Text = TextBox2.Text
            Hotel_Book.searchCustomer(con)
            Me.Close()
        Else
            MessageBox.Show("Select a Customer!")
        End If
    End Sub
End Class