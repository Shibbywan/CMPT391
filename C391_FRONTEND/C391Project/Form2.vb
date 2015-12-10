Imports System.Data.SqlClient

Public Class rooms

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Hotel_Book.LoadRoomTypes(con)
        Me.Close()
    End Sub


    Private Sub rooms_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Me.StartPosition = FormStartPosition.CenterScreen
        LoadRoomTypes(con)
    End Sub

    Private Sub LoadRoomTypes(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Dim cid = Hotel_Book.getCID

        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Type_ID, Name From Room_Type", connection)
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadReservations")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select Type_ID, Name From Room_Type"
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

    Private Sub addRoomType(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("addService2")
            command.Connection = connection
            command.Transaction = transaction
            Dim id = TextBox2.Text
            Dim name = TextBox3.Text
            Dim desc = TextBox1.Text
            Dim price = TextBox4.Text
            Try

                command.CommandText = _
                "Insert into Room_Type (Type_ID,Name,Description,Price) VALUES ('" & id & "','" & name & "','" & desc & "','" & price & "')"
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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        addRoomType(con)
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        LoadRoomTypes(con)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadRoomTypes(con)
    End Sub
End Class