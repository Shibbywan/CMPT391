Imports System.Data.SqlClient

Public Class services

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub LoadServices(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Dim resid = Hotel_Book.getRESID
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Name, Price from Services, Service_Type, Reservation Where Services.S_ID = Service_Type.S_ID and Reservation.C_ID = Services.C_ID and Reservation.R_ID ='" & TextBox1.Text & "' and Reservation.RES_ID ='" & resid & "' and Services.RES_ID = Reservation.RES_ID", connection)
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("Load Customers")
            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select Name, Price from Services, Service_Type, Reservation Where Services.S_ID = Service_Type.S_ID and Reservation.C_ID = Services.C_ID and Reservation.R_ID ='" & TextBox1.Text & "' and Reservation.RES_ID ='" & resid & "' and Services.RES_ID = Reservation.RES_ID"
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

    Private Sub GetCustomerInfo(ByVal connectionString As String)
        Dim resid = Hotel_Book.getRESID
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction


            transaction = connection.BeginTransaction("GetCustomerInfo")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                    "Select Reservation.R_ID From Customer, Reservation, Rooms Where Reservation.R_ID = Rooms.R_ID and Reservation.C_ID = Customer.C_ID and Reservation.RES_ID ='" & resid & "'"
                TextBox1.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select Customer.C_ID From Customer,Reservation Where Reservation.C_ID = Customer.C_ID and Reservation.RES_ID ='" & resid & "'"
                TextBox2.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select Name From Customer,Reservation,Person Where Person.P_ID = Reservation.C_ID and Customer.C_ID = Reservation.C_ID and RES_ID ='" & resid & "'"
                TextBox3.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select SUM(Price) from Services, Service_Type, Reservation Where Services.S_ID = Service_Type.S_ID and Reservation.C_ID = Services.C_ID and Reservation.R_ID ='" & TextBox1.Text & "' and Reservation.RES_ID ='" & resid & "' and Services.RES_ID = Reservation.RES_ID"
                TextBox4.Text = Convert.ToString(command.ExecuteScalar())
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

    Private Sub services_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Me.StartPosition = FormStartPosition.CenterScreen
        GetCustomerInfo(con)
        LoadServices(con)
    End Sub


End Class