Imports System.Data.SqlClient

Public Class reservations

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub reservations_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Me.StartPosition = FormStartPosition.CenterScreen
        LoadReservations(con)
    End Sub

    Private Sub LoadReservations(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Dim cid = Hotel_Book.getCID

        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Invoice.RES_ID as 'Reservation ID', Check_In_Date as 'Check In', Inv_Date as 'Check Out'  From Reservation, Invoice Where Reservation.RES_ID = Invoice.RES_ID and Invoice.C_ID =" & cid, connection)
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadReservations")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select Invoice.RES_ID as 'Reservation ID', Check_In_Date as 'Check In', Inv_Date as 'Check Out'  From Reservation, Invoice Where Reservation.RES_ID = Invoice.RES_ID and Invoice.C_ID =" & cid
                command.ExecuteNonQuery()

                command.CommandText = _
                 "Select C_ID From Customer Where C_ID = '" & cid & "'"
                TextBox2.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                  "Select Name From Customer, Person Where Customer.C_ID = Person.P_ID and C_ID = '" & cid & "'"
                TextBox3.Text = Convert.ToString(command.ExecuteScalar())
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