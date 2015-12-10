Imports System.Data.SqlClient
Imports System.Timers

Public Class Hotel_Book
    Private Timer As New System.Timers.Timer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterScreen
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadRooms(con)
        LoadCustomers(con)
        ReservationCount(con)
        PersonCount(con)
        LoadOccupiedRooms(con)
        LoadOccupiedRooms2(con)
        LoadAllRooms(con)
        LoadAllRooms2(con)
        LoadEmployees(con)
        LoadEmployees2(con)
        LoadServices(con)
        LoadServices2(con)
        LoadPayments(con)
        GetServiceInfo(con)
        LoadRoomTypes(con)
        LoadPaymentMethods(con)
        setPrice(con)
        LoginForm1.Show()
        LoginForm1.TopMost = True
    End Sub

    Private Sub LoadCustomers(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("SELECT C_ID as 'Customer ID', Name FROM Customer, Person WHERE Customer.C_ID = Person.P_ID", connection)
            adapter.Fill(ds)
            customerTable.DataSource = ds.Tables(0)
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

    Private Sub LoadEmployees(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("SELECT E_ID as 'Employee ID', Name FROM Employee, Person WHERE Employee.E_ID = Person.P_ID", connection)
            adapter.Fill(ds)
            employeeTable.DataSource = ds.Tables(0)

            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("Load Employees")
            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "SELECT E_ID as 'Employee ID', Name FROM Employee, Person WHERE Employee.E_ID = Person.P_ID"
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

    Private Sub LoadRooms(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Dim dt1 As String = checkinDate.Value.Year & "-" & checkinDate.Value.Month & "-" & checkinDate.Value.Day
        Dim time1 As DateTime = DateTime.Parse(dt1)
        Dim dt2 As String = checkoutDate.Value.Year & "-" & checkoutDate.Value.Month & "-" & checkoutDate.Value.Day
        Dim time2 As DateTime = DateTime.Parse(dt2)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("SELECT Distinct Rooms.[R_ID] as 'Room ID', [Floor], [Name], [Price] FROM [Rooms], [Room_Type], [Reservation] WHERE Rooms.TYPE_ID = Room_Type.TYPE_ID  and Rooms.R_ID NOT IN ( SELECT ROOMS.[R_ID] FROM [Rooms], [Room_Type], [Reservation] WHERE Rooms.Type_ID = Room_Type.TYPE_ID and Reservation.R_ID = Rooms.R_ID and Reservation.Check_Out_Date > " & "'" & dt1 & "'" & " and Reservation.Check_In_Date < " & "'" & dt2 & "')", connection)
            adapter.Fill(ds)
            availableRoomTable.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadRooms")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                "SELECT Distinct Rooms.[R_ID] as 'Room ID', [Floor], [Name], [Price] FROM [Rooms], [Room_Type], [Reservation] WHERE Rooms.TYPE_ID = Room_Type.TYPE_ID  and Rooms.R_ID NOT IN ( SELECT ROOMS.[R_ID] FROM [Rooms], [Room_Type], [Reservation] WHERE Rooms.Type_ID = Room_Type.TYPE_ID and Reservation.R_ID = Rooms.R_ID and Reservation.Check_Out_Date > " & "'" & dt1 & "'" & " and Reservation.Check_In_Date < " & "'" & dt2 & "')"
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

    Private Sub LoadEmployees2(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("SELECT E_ID as 'Employee ID', Name FROM Employee, Person WHERE Employee.E_ID = Person.P_ID", connection)
            adapter.Fill(ds)
            employeeTable2.DataSource = ds.Tables(0)

            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("Load Employees")
            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "SELECT E_ID as 'Employee ID', Name FROM Employee, Person WHERE Employee.E_ID = Person.P_ID"
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

    Private Sub refreshButton_Click(sender As Object, e As EventArgs) Handles refreshButton.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadCustomers(con)
    End Sub

    Private Sub committedToDB()
        Dim tmr As New System.Timers.Timer()
        tmr.Interval = 5000
        tmr.Enabled = True
        tmr.Start()
        Panel3.Visible = True
        AddHandler tmr.Elapsed, AddressOf OnTimedEvent
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles confirmButton1.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (nameField.TextLength > 0 And cidField.TextLength > 0 And countryField.TextLength > 0 And provinceField.TextLength > 0 And addressField.TextLength > 0 And postalCodeField.TextLength > 0 And cityField.TextLength > 0 And phoneField.TextLength > 0 And TextBox9.TextLength > 0 And TextBox10.TextLength > 0) Then
            If (IsInDatagridview(cidField.Text, 0, customerTable) = False) Then
                ExecuteSqlTransaction(con)
            Else
                ReturningCustomer(con)
            End If
            nameField.Clear()
            countryField.Clear()
            addressField.Clear()
            cityField.Clear()
            phoneField.Clear()
            TextBox10.Clear()
            provinceField.Clear()
            postalCodeField.Clear()
            availableRoomTable.ClearSelection()
            TextBox9.Clear()
            TextBox11.Clear()
            TextBox12.Clear()
            LoadRooms(con)
        Else
            MessageBox.Show("Missing Field(s)!")
        End If
    End Sub

    Private Delegate Sub CloseFormCallback()

    Private Sub ClosePanel()
        If InvokeRequired Then
            Dim d As New CloseFormCallback(AddressOf ClosePanel)
            Invoke(d, Nothing)
        Else
            Panel3.Visible = False
        End If
    End Sub

    Private Sub OnTimedEvent(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        ClosePanel()
    End Sub

    Private Sub ExecuteSqlTransaction(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("AvailableRooms")
            command.Connection = connection
            command.Transaction = transaction
            Dim name As String = nameField.Text
            Dim cid As String = cidField.Text
            Dim country As String = countryField.Text
            Dim address As String = addressField.Text
            Dim city As String = cityField.Text
            Dim phone As String = phoneField.Text
            Dim post As String = postalCodeField.Text
            Dim prov As String = provinceField.Text
            Dim i = availableRoomTable.CurrentRow.Index
            Dim RID = availableRoomTable.Item(0, i).Value
            Dim dt1 As String = checkinDate.Value.Year & "-" & checkinDate.Value.Month & "-" & checkinDate.Value.Day
            Dim time1 As DateTime = DateTime.Parse(dt1)
            Dim dt2 As String = checkoutDate.Value.Year & "-" & checkoutDate.Value.Month & "-" & checkoutDate.Value.Day
            Dim time2 As DateTime = DateTime.Parse(dt2)
            Try

                command.CommandText = _
                "Insert into Person (Name, P_ID) VALUES (" & "'" & name & "'," & "'" & cid & "'" & ")"
                command.ExecuteNonQuery()

                command.CommandText = _
                "Insert into Customer (C_Address, C_Postal_Code, C_Country, C_ID, C_City, C_Phone,C_Province) VALUES (" & "'" & address & "'" & "," & "'" & post & "','" & country & "'" & "," & "'" & cid & "'" & "," & "'" & city & "'" & "," & "'" & phone & "','" & prov & "')"
                command.ExecuteNonQuery()

                command.CommandText = _
                  "Insert into Reservation (C_ID, RES_ID, R_ID, Check_In_Date, Check_Out_Date) VALUES (" & "'" & cid & "'," & "'" & reservationCountLabel.Text.ToString & "'," & "'" & RID & "'," & "'" & time1 & "'," & "'" & time2 & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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
            ReservationCount(con)
            PersonCount(con)
        End Using
    End Sub

    Private Sub ReservationCount(ByVal connectionString As String)
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            Dim count As Int16
            transaction = connection.BeginTransaction("ReservationCount")
            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                 "SELECT COUNT(*) FROM Reservation"
                command.ExecuteNonQuery()
                count = Convert.ToInt16(command.ExecuteScalar())
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
            count = count + 1
            reservationCountLabel.Text = count.ToString
        End Using
    End Sub

    Private Sub PersonCount(ByVal connectionString As String)
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            Dim count As Int16
            transaction = connection.BeginTransaction("CustomerCount")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                 "SELECT COUNT(*) FROM Person"
                command.ExecuteNonQuery()
                count = Convert.ToInt16(command.ExecuteScalar())
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
            count = count + 1
            cidField.Text = count.ToString
            TextBox51.Text = count.ToString
        End Using
    End Sub

    Private Sub LoadOccupiedRooms(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Name, RES_ID as 'Reservation ID', R_ID as 'Room ID', Check_In_Date as 'Check In', Check_Out_Date as 'Check Out' From Reservation, Person Where Check_Out_Date >  Check_In_Date and Person.P_ID = Reservation.C_ID and Reservation.RES_ID NOT IN (Select RES_ID from Invoice)", connection)
            adapter.Fill(ds)
            occupiedRoomsTable.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadOccupiedRooms")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                "Select Name, RES_ID as 'Reservation ID', R_ID as 'Room ID', Check_In_Date as 'Check In', Check_Out_Date as 'Check Out' From Reservation, Person Where Check_Out_Date >  Check_In_Date and Person.P_ID = Reservation.C_ID and Reservation.RES_ID NOT IN (Select RES_ID from Invoice)"
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

    Private Sub LoadOccupiedRooms2(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Name, RES_ID as 'Reservation ID', R_ID as 'Room ID', Check_In_Date as 'Check In', Check_Out_Date as 'Check Out' From Reservation, Person Where Check_Out_Date >  Check_In_Date and Person.P_ID = Reservation.C_ID and Reservation.RES_ID NOT IN (Select RES_ID from Invoice)", connection)
            adapter.Fill(ds)
            occupiedRoomsTable2.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadOccupiedRooms")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                "Select Name, RES_ID as 'Reservation ID', R_ID as 'Room ID', Check_In_Date as 'Check In', Check_Out_Date as 'Check Out' From Reservation, Person Where Check_Out_Date >  Check_In_Date and Person.P_ID = Reservation.C_ID and Reservation.RES_ID NOT IN (Select RES_ID from Invoice)"
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

    Private Sub LoadAllRooms(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select R_ID as 'Room ID', Price, Name From Rooms, Room_Type Where Rooms.TYPE_ID = Room_Type.TYPE_ID", connection)
            adapter.Fill(ds)
            allRoomsTable.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadAllRooms")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select R_ID as 'Room ID', Price, Name From Rooms, Room_Type Where Rooms.TYPE_ID = Room_Type.TYPE_ID"
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

    Private Sub dataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles availableRoomTable.CellContentClick
        Dim dgv As DataGridView = TryCast(sender, DataGridView)
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If dgv Is Nothing Then
            Return
        End If
        Dim i = availableRoomTable.CurrentRow.Index
        Dim RID = availableRoomTable.Item(0, i).Value
        TextBox9.Text = RID.ToString
        If dgv.CurrentRow.Selected Then
        End If
        GetDescription(con)
    End Sub

    Private Sub GetDescription(ByVal connectionString As String)
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            Dim RID = TextBox9.Text.ToString


            transaction = connection.BeginTransaction("GetDescription")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                 "SELECT Description FROM Rooms Where Rooms.R_ID = " & "'" & RID & "'"
                command.ExecuteNonQuery()
                TextBox12.Text = command.ExecuteScalar()

                command.CommandText = _
                 "SELECT Room_Type.Name FROM Room_Type, Rooms WHERE Rooms.TYPE_ID = Room_Type.TYPE_ID and Rooms.R_ID =" & "'" & RID & "'"
                command.ExecuteNonQuery()
                TextBox11.Text = command.ExecuteScalar()
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

    Private Sub ExecuteCheckout(ByVal connectionString As String)
        Dim ds As New DataSet
        Dim name = TextBox1.Text
        Dim cid As Int16
        Dim inv As Int16
        Dim M_ID As Int16
        Dim cost = TextBox3.Text
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            Dim RESID = TextBox6.Text.ToString

            transaction = connection.BeginTransaction("ExecuteCheckout")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                    "Select C_ID From Person,Customer Where Name = '" & name & "' and Person.P_ID = Customer.C_ID"
                cid = Convert.ToInt16(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select Count(*) From Invoice"
                inv = Convert.ToInt16(command.ExecuteScalar()) + 1
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select M_ID From Payment_Method Where Method='" & ComboBox1.GetItemText(ComboBox1.SelectedItem) & "'"
                M_ID = Convert.ToInt16(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Insert Into Invoice(C_ID, INV_ID, M_ID, Services_Cost, Inv_Date, Total_Cost,RES_ID, E_ID) VALUES ('" & cid & "','" & inv & "','" & M_ID & "','" & TextBox4.Text & "',GETDATE(),'" & cost & "','" & TextBox6.Text & "','" & TextBox55.Text & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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

    Private Sub GetCheckoutInfo(ByVal connectionString As String)
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            Dim RID = TextBox9.Text.ToString

            transaction = connection.BeginTransaction("GetCheckoutInfo")
            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                 "SELECT Name FROM Person, Customer Where Rooms.R_ID = " & "'" & RID & "'"
                command.ExecuteNonQuery()
                TextBox12.Text = command.ExecuteScalar()

                command.CommandText = _
                 "SELECT Room_Type.Name FROM Room_Type, Rooms WHERE Rooms.TYPE_ID = Room_Type.TYPE_ID and Rooms.R_ID =" & "'" & RID & "'"
                command.ExecuteNonQuery()
                TextBox11.Text = command.ExecuteScalar()
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



    Private Sub LoadServices2(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select S_ID as 'Service ID', Name, Price From Service_Type", connection)
            adapter.Fill(ds)
            serviceTable.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadServices2")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select S_ID as 'Service ID', Name, Price From Service_Type"
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

    Private Sub LoadAllRooms2(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select R_ID as 'Room ID', Price, Name From Rooms, Room_Type Where Rooms.TYPE_ID = Room_Type.TYPE_ID", connection)
            adapter.Fill(ds)
            allRoomsTable2.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadAllRooms")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select R_ID as 'Room ID', Price, Name From Rooms, Room_Type Where Rooms.TYPE_ID = Room_Type.TYPE_ID"
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


    Private Sub GetCheckoutCost(ByVal connectionString As String)
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            Dim serv As Single
            Dim room As Single
            Dim cost As Single


            transaction = connection.BeginTransaction("GetCheckoutCost")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                    "Select SUM(Price) from Services, Service_Type, Reservation Where Services.S_ID = Service_Type.S_ID and Reservation.C_ID = Services.C_ID and Reservation.R_ID ='" & TextBox59.Text & "' and Reservation.RES_ID ='" & TextBox6.Text & "' and Services.RES_ID = Reservation.RES_ID"
                If (IsDBNull(command.ExecuteScalar)) Then
                    serv = 0
                Else
                    serv = Convert.ToSingle(command.ExecuteScalar())
                End If

                command.ExecuteNonQuery()



                command.CommandText = _
                    "Select Price From Customer, Rooms, Room_Type, Reservation Where Customer.C_ID = Reservation.C_ID and Rooms.R_ID = Reservation.R_ID and Rooms.Type_ID = Room_Type.Type_ID and Reservation.RES_ID ='" & TextBox6.Text & "'"
                room = Convert.ToSingle(command.ExecuteScalar())
                command.ExecuteNonQuery()

                cost = room + serv
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
            TextBox4.Text = serv.ToString
            TextBox3.Text = cost.ToString
        End Using

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles clearButton1.Click
        postalCodeField.Clear()
        provinceField.Clear()
        nameField.Clear()
        countryField.Clear()
        addressField.Clear()
        cityField.Clear()
        phoneField.Clear()
        TextBox10.Clear()
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles checkinDate.ValueChanged
        Dim dt1 As String = checkinDate.Value.Year & "-" & checkinDate.Value.Month & "-" & checkinDate.Value.Day
        Dim time1 As DateTime = DateTime.Parse(dt1)
        Dim dt2 As String = checkoutDate.Value.Year & "-" & checkoutDate.Value.Month & "-" & checkoutDate.Value.Day
        Dim time2 As DateTime = DateTime.Parse(dt2)
        If (time1 > time2) Then
            MessageBox.Show("Invalid Date Range!")
            checkoutDate.Value = time1
        Else
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            LoadRooms(con)
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles checkoutDate.ValueChanged
        Dim dt1 As String = checkinDate.Value.Year & "-" & checkinDate.Value.Month & "-" & checkinDate.Value.Day
        Dim time1 As DateTime = DateTime.Parse(dt1)
        Dim dt2 As String = checkoutDate.Value.Year & "-" & checkoutDate.Value.Month & "-" & checkoutDate.Value.Day
        Dim time2 As DateTime = DateTime.Parse(dt2)
        If (time1 > time2) Then
            MessageBox.Show("Invalid Date Range!")
            checkoutDate.Value = time1
        Else
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            LoadRooms(con)
        End If

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles occupiedRoomsTable.CellContentClick
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Dim i = occupiedRoomsTable.CurrentRow.Index
        Dim RID = occupiedRoomsTable.Item(2, i).Value
        Dim RESID = occupiedRoomsTable.Item(1, i).Value
        Dim name = occupiedRoomsTable.Item(0, i).Value
        TextBox59.Text = RID.ToString
        TextBox1.Text = name.ToString
        TextBox6.Text = RESID.ToString
        GetCheckoutCost(con)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles confirmButton2.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (TextBox1.TextLength > 0 And TextBox6.TextLength > 0 And TextBox3.Text <> "0") Then
            ExecuteCheckout(con)
            LoadOccupiedRooms(con)
            TextBox1.Clear()
            TextBox6.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox59.Clear()
        Else
            MessageBox.Show("Invalid Field(s)!")
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles refreshButton1.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadOccupiedRooms(con)
        LoadPayments(con)
    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If (ComboBox1.GetItemText(ComboBox1.SelectedItem).Contains("Visa")) Then
            PictureBox1.Image = ImageList2.Images(2)
        End If
        If (ComboBox1.GetItemText(ComboBox1.SelectedItem).Contains("Mastercard")) Then
            PictureBox1.Image = ImageList2.Images(1)
        End If
        If (ComboBox1.GetItemText(ComboBox1.SelectedItem).Contains("Debit")) Then
            PictureBox1.Image = ImageList2.Images(3)
        End If
        If (ComboBox1.GetItemText(ComboBox1.SelectedItem).Contains("Cash")) Then
            PictureBox1.Image = ImageList2.Images(4)
        End If
        If (ComboBox1.GetItemText(ComboBox1.SelectedItem).Contains("Airmiles")) Then
            PictureBox1.Image = ImageList2.Images(5)
        End If

    End Sub


    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If (TextBox6.TextLength > 0) Then
            services.Show()
        Else
            MessageBox.Show("Enter a Reservation ID!")
        End If
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        If (TextBox2.TextLength > 0) Then
            reservations.Show()
        Else
            MessageBox.Show("Select a Customer!")
        End If
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs)
        services.Show()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        rooms.Show()
    End Sub



    Private Sub cancelButton1_Click(sender As Object, e As EventArgs) Handles cancelButton1.Click
        TextBox6.Clear()
        TextBox1.Clear()
        TextBox4.Clear()
        TextBox3.Clear()
    End Sub

    Public Function getCID() As String
        getCID = TextBox2.Text
    End Function

    Public Function getRESID() As String
        getRESID = TextBox6.Text
    End Function


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub GetCustomerInfo(ByVal connectionString As String)
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
                    "Select C_Country From Customer Where C_ID ='" & TextBox2.Text & "'"
                TextBox7.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select C_Province From Customer Where C_ID ='" & TextBox2.Text & "'"
                TextBox25.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select C_City From Customer Where C_ID ='" & TextBox2.Text & "'"
                TextBox8.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select C_Postal_Code From Customer Where C_ID ='" & TextBox2.Text & "'"
                TextBox13.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select C_Phone From Customer Where C_ID ='" & TextBox2.Text & "'"
                TextBox14.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select R_ID From Customer, Reservation Where Customer.C_ID = Reservation.C_ID and Customer.C_ID = " & "'" & TextBox2.Text & "' and Reservation.Check_Out_Date > GETDATE()"
                If (Convert.ToString(command.ExecuteScalar()) <> "") Then
                    TextBox16.Text = "Room " & Convert.ToString(command.ExecuteScalar())
                Else
                    TextBox16.Text = "N/A"
                End If
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select Inv_Date From Invoice, Customer where Invoice.C_ID = Customer.C_ID and Customer.C_ID ='" & TextBox2.Text & "' and Invoice.Inv_Date < GETDATE()"
                If (Convert.ToString(command.ExecuteScalar()) <> "") Then
                    TextBox15.Text = Convert.ToString(command.ExecuteScalar())
                Else
                    TextBox15.Text = "N/A"
                End If
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

    Private Sub LoadPayments(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Distinct Method From Payment_Method", connection)
            adapter.Fill(ds)
            ComboBox1.DataSource = ds.Tables(0)
            ComboBox1.DisplayMember = "Method"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadAllRooms")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select Distinct Method From Payment_Method"
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

    Private Sub LoadServices(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Distinct Name From Service_Type", connection)
            adapter.Fill(ds)
            ComboBox2.DataSource = ds.Tables(0)
            ComboBox2.DisplayMember = "Name"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadServices")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select Distinct Name From Service_Type"
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

    Private Sub customerTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles customerTable.CellContentClick
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Dim i = customerTable.CurrentRow.Index
        Dim cid = customerTable.Item(1, i).Value
        Dim name = customerTable.Item(0, i).Value
        TextBox2.Text = name
        TextBox5.Text = cid
        GetCustomerInfo(con)
    End Sub



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadOccupiedRooms(con)
        LoadPayments(con)
        TextBox6.Clear()
        TextBox4.Clear()
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        If (TextBox1.TextLength <> 0) Then
            Using connection As New SqlConnection(con)
                connection.Open()
                adapter = New SqlDataAdapter("Select Name, RES_ID as 'Reservation ID', R_ID as 'Room ID', Check_In_Date as 'Check In', Check_Out_Date as 'Check Out' From Reservation, Person Where Check_Out_Date >  Check_In_Date and Person.P_ID = Reservation.C_ID and Name LIKE '%" & TextBox1.Text & "% ' and Reservation.RES_ID NOT IN (Select RES_ID from Invoice)", connection)
                adapter.Fill(ds)
                occupiedRoomsTable.DataSource = ds.Tables(0)
                Dim command As SqlCommand = connection.CreateCommand()
                Dim transaction As SqlTransaction

                transaction = connection.BeginTransaction("LoadOccupiedRooms")

                command.Connection = connection
                command.Transaction = transaction


                Try
                    command.CommandText = _
                    "Select Name, RES_ID as 'Reservation ID', R_ID as 'Room ID', Check_In_Date as 'Check In', Check_Out_Date as 'Check Out' From Reservation, Person Where Check_Out_Date >  Check_In_Date and Person.P_ID = Reservation.C_ID and Name LIKE '%" & TextBox1.Text & "% ' and Reservation.RES_ID NOT IN (Select RES_ID from Invoice)"
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
        Else
            LoadOccupiedRooms(con)
        End If
    End Sub


    Private Sub GetEmployeeInfo(ByVal connectionString As String)
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction


            transaction = connection.BeginTransaction("GetEmployeeInfo")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                    "Select E_Position From Employee Where E_ID ='" & TextBox24.Text & "'"
                TextBox22.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select E_Country From Employee Where E_ID ='" & TextBox24.Text & "'"
                TextBox21.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select E_Province From Employee Where E_ID ='" & TextBox24.Text & "'"
                TextBox20.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select E_City From Employee Where E_ID ='" & TextBox24.Text & "'"
                TextBox19.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select E_Postal_Code From Employee Where E_ID ='" & TextBox24.Text & "'"
                TextBox27.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select E_Phone From Employee Where E_ID ='" & TextBox24.Text & "'"
                TextBox18.Text = Convert.ToString(command.ExecuteScalar())
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

    Private Sub GetRoomInfo(ByVal connectionString As String)
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            Dim RID = TextBox32.Text.ToString


            transaction = connection.BeginTransaction("GetRoomInfo")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                 "SELECT Description FROM Rooms Where Rooms.R_ID = " & "'" & RID & "'"
                command.ExecuteNonQuery()
                TextBox31.Text = command.ExecuteScalar()

                command.CommandText = _
                 "SELECT Rooms.Floor FROM Room_Type, Rooms WHERE Rooms.TYPE_ID = Room_Type.TYPE_ID and Rooms.R_ID =" & "'" & RID & "'"
                command.ExecuteNonQuery()
                TextBox30.Text = command.ExecuteScalar()

                command.CommandText = _
                 "SELECT Rooms.Num_Of_Beds FROM Room_Type, Rooms WHERE Rooms.TYPE_ID = Room_Type.TYPE_ID and Rooms.R_ID =" & "'" & RID & "'"
                command.ExecuteNonQuery()
                TextBox29.Text = command.ExecuteScalar()

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


    Private Sub GetServiceInfo(ByVal connectionString As String)
        Dim ds As New DataSet
        Dim serv = ComboBox2.GetItemText(ComboBox2.SelectedItem)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction


            transaction = connection.BeginTransaction("GetServiceInfo")

            command.Connection = connection
            command.Transaction = transaction


            Try
                command.CommandText = _
                    "Select Price From Service_Type Where Name = '" & serv & "'"
                TextBox36.Text = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                    "Select Department From Service_Type Where Name = '" & serv & "'"
                TextBox35.Text = Convert.ToString(command.ExecuteScalar())
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

    Private Sub addService(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("addService")
            command.Connection = connection
            command.Transaction = transaction
            Dim rid = TextBox40.Text
            Dim name = TextBox39.Text
            Dim serv = ComboBox2.GetItemText(ComboBox2.SelectedItem)
            Try

                command.CommandText = _
                "Select S_ID From Service_Type Where Name = '" & serv & "'"
                Dim sid = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                "Select C_ID From Customer,Person Where Customer.C_ID = Person.P_ID and Name = '" & name & "'"
                Dim cid = Convert.ToString(command.ExecuteScalar())
                command.ExecuteNonQuery()

                command.CommandText = _
                "Insert into Services (S_ID, C_ID, Date, E_ID, RES_ID) VALUES ('" & sid & "','" & cid & "', GETDATE(),'" & TextBox28.Text & "','" & TextBox58.Text & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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

    Private Sub addService2(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("addService2")
            command.Connection = connection
            command.Transaction = transaction
            Dim sid = TextBox41.Text
            Dim name = TextBox38.Text
            Dim price = TextBox34.Text
            Dim desc = TextBox37.Text
            Try

                command.CommandText = _
                "Insert into Service_Type (S_ID, Name, Department, Price) VALUES ('" & sid & "','" & name & "','" & desc & "','" & price & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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

    Public Sub LoadRoomTypes(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select Distinct Name From Room_Type", connection)
            adapter.Fill(ds)
            ComboBox3.DataSource = ds.Tables(0)
            ComboBox3.DisplayMember = "Name"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadRoomTypes")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select Distinct Name From Room_Type"
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

    Private Sub addRoom(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("addRoom")
            command.Connection = connection
            command.Transaction = transaction
            Dim rid = TextBox45.Text
            Dim desc = TextBox43.Text
            Dim floor = TextBox42.Text
            Dim numbeds = TextBox44.Text
            Dim roomtype = ComboBox3.GetItemText(ComboBox3.SelectedItem)
            Try

                command.CommandText = _
                "Select Room_Type.Type_ID From Room_Type Where Room_Type.Name ='" & roomtype.Trim & "'"
                command.ExecuteNonQuery()
                Dim typeid = Convert.ToInt16(command.ExecuteScalar())

                command.CommandText = _
                "Insert into Rooms (R_ID, Description, Num_of_Beds,Floor, Type_ID) VALUES ('" & rid & "','" & desc & "','" & numbeds & "','" & floor & "','" & typeid & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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

    Private Sub addEmployee(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("addEmployee")
            command.Connection = connection
            command.Transaction = transaction
            Dim eid = TextBox51.Text
            Dim name = TextBox48.Text
            Dim pos = TextBox50.Text
            Dim coun = TextBox47.Text
            Dim prov = TextBox49.Text
            Dim city = TextBox52.Text
            Dim addr = TextBox53.Text
            Dim phone = TextBox54.Text
            Dim post = TextBox26.Text

            Try

                command.CommandText = _
                "Insert into Person (Name, P_ID) VALUES ('" & name & "','" & eid & "')"
                command.ExecuteNonQuery()

                command.CommandText = _
                "Insert into Employee(E_ID, E_Position, E_Address, E_Province, E_Country, E_City, E_Postal_Code, E_Phone) VALUES ('" & eid & "','" & pos & "','" & addr & "','" & prov & "','" & coun & "','" & city & "','" & post & "','" & phone & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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

            PersonCount(con)
        End Using
    End Sub

    Private Sub setPrice(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("setPrice")
            command.Connection = connection
            command.Transaction = transaction
            Dim roomtype = ComboBox3.GetItemText(ComboBox3.SelectedItem)

            Try

                command.CommandText = _
                "Select Price From Room_Type Where Name='" & roomtype & "'"
                TextBox46.Text = Convert.ToInt16(command.ExecuteScalar())
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

            PersonCount(con)
        End Using

    End Sub

    Private Sub LoadPaymentMethods(ByVal connectionString As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter("Select M_ID as 'Method ID', Method From Payment_Method", connection)
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("LoadPaymentMethods")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                 "Select M_ID as 'Method ID', Method From Payment_Method"
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

    Private Sub addPayment(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("addPayment")
            command.Connection = connection
            command.Transaction = transaction
            Dim mid = TextBox57.Text
            Dim name = TextBox56.Text
            Try
                command.CommandText = _
                "Insert into Payment_Method (M_ID, Method) VALUES ('" & mid & "','" & name & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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

    Public Sub searchCustomer(ByVal connectionString As String)
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (cidField.TextLength <> 0) Then
            Using connection As New SqlConnection(con)
                connection.Open()
                Dim command As SqlCommand = connection.CreateCommand()
                Dim transaction As SqlTransaction

                transaction = connection.BeginTransaction("LoadOccupiedRooms")
                command.Connection = connection
                command.Transaction = transaction


                Try
                    command.CommandText = _
                    "Select Name From Person Where P_ID = " & cidField.Text
                    nameField.Text = command.ExecuteScalar()
                    command.ExecuteNonQuery()

                    command.CommandText = _
                    "Select C_Country From Customer Where C_ID = " & cidField.Text
                    countryField.Text = command.ExecuteScalar()
                    command.ExecuteNonQuery()

                    command.CommandText = _
                    "Select C_Province From Customer Where C_ID = " & cidField.Text
                    provinceField.Text = command.ExecuteScalar()
                    command.ExecuteNonQuery()

                    command.CommandText = _
                    "Select C_Address From Customer Where C_ID = " & cidField.Text
                    addressField.Text = command.ExecuteScalar()
                    command.ExecuteNonQuery()

                    command.CommandText = _
                    "Select C_Postal_Code From Customer Where C_ID = " & cidField.Text
                    postalCodeField.Text = command.ExecuteScalar()
                    command.ExecuteNonQuery()

                    command.CommandText = _
                    "Select C_City From Customer Where C_ID = " & cidField.Text
                    cityField.Text = command.ExecuteScalar()
                    command.ExecuteNonQuery()

                    command.CommandText = _
                    "Select C_Phone From Customer Where C_ID = " & cidField.Text
                    phoneField.Text = command.ExecuteScalar()
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
        End If
    End Sub

    Private Sub employeeTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles employeeTable.CellContentClick
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Dim i = employeeTable.CurrentRow.Index
        Dim name = employeeTable.Item(1, i).Value
        Dim eid = employeeTable.Item(0, i).Value
        TextBox23.Text = name
        TextBox24.Text = eid
        GetEmployeeInfo(con)
    End Sub

    Private Sub allRoomsTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles allRoomsTable.CellContentClick
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Dim i = allRoomsTable.CurrentRow.Index
        Dim name = allRoomsTable.Item(2, i).Value
        Dim price = allRoomsTable.Item(1, i).Value
        Dim Type = allRoomsTable.Item(0, i).Value
        TextBox33.Text = name
        TextBox32.Text = Type
        TextBox17.Text = price
        GetRoomInfo(con)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadAllRooms(con)
        TextBox33.Clear()
        TextBox32.Clear()
        TextBox17.Clear()
        TextBox29.Clear()
        TextBox30.Clear()
        TextBox31.Clear()
    End Sub


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadOccupiedRooms2(con)
        LoadServices2(con)
    End Sub

    Private Sub occupiedRoomsTable2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles occupiedRoomsTable2.CellContentClick
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        Dim i = occupiedRoomsTable2.CurrentRow.Index
        Dim name = occupiedRoomsTable2.Item(0, i).Value
        Dim resid = occupiedRoomsTable2.Item(1, i).Value
        Dim room = occupiedRoomsTable2.Item(2, i).Value
        TextBox40.Text = room
        TextBox58.Text = resid
        TextBox39.Text = name
        GetServiceInfo(con)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        addService(con)
        TextBox40.Clear()
        TextBox58.Clear()
        TextBox39.Clear()
        TextBox36.Clear()
        TextBox35.Clear()
        TextBox28.Clear()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        GetServiceInfo(con)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox41.Clear()
        TextBox38.Clear()
        TextBox34.Clear()
        TextBox37.Clear()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadServices2(con)
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (TextBox41.TextLength > 0 And TextBox38.TextLength > 0 And TextBox34.TextLength > 0 And TextBox37.TextLength > 0 And IsInDatagridview(TextBox41.Text, 0, serviceTable) = False) Then
            addService2(con)
            LoadServices2(con)
            TextBox41.Clear()
            TextBox38.Clear()
            TextBox34.Clear()
            TextBox37.Clear()
        Else
            MessageBox.Show("Invalid Field(s)!")
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)
        Dim i = serviceTable.CurrentRow.Index
        Dim sid = serviceTable.Item(0, i).Value

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (TextBox45.TextLength > 0 And TextBox44.TextLength > 0 And TextBox42.TextLength > 0 And TextBox43.TextLength > 0 And IsInDatagridview(TextBox45.Text, 0, allRoomsTable2) = False) Then
            addRoom(con)
            LoadAllRooms2(con)
            TextBox45.Clear()
            TextBox44.Clear()
            TextBox42.Clear()
            TextBox43.Clear()
        Else
            MessageBox.Show("Invalid Field(s)!")
        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (TextBox51.TextLength > 0 And TextBox48.TextLength > 0 And TextBox50.TextLength > 0 And TextBox47.TextLength > 0 And TextBox49.TextLength > 0 And TextBox52.TextLength > 0 And TextBox53.TextLength > 0 And TextBox54.TextLength > 0 And TextBox26.TextLength > 0 And IsInDatagridview(TextBox51.Text, 0, employeeTable2) = False) Then
            addEmployee(con)
            PersonCount(con)
            TextBox48.Clear()
            TextBox50.Clear()
            TextBox47.Clear()
            TextBox49.Clear()
            TextBox52.Clear()
            TextBox53.Clear()
            TextBox54.Clear()
            TextBox26.Clear()
            LoadEmployees2(con)
        Else
            MessageBox.Show("Invalid Field(s)!")
        End If
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        PersonCount(con)
        TextBox48.Clear()
        TextBox50.Clear()
        TextBox47.Clear()
        TextBox49.Clear()
        TextBox52.Clear()
        TextBox53.Clear()
        TextBox54.Clear()
        TextBox26.Clear()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadEmployees2(con)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        setPrice(con)
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        If (TextBox56.TextLength > 0 And TextBox57.TextLength > 0 And IsInDatagridview(TextBox57.Text, 0, DataGridView1) = False) Then
            addPayment(con)
            LoadPaymentMethods(con)
            TextBox56.Clear()
            TextBox57.Clear()
        Else
            MessageBox.Show("Invalid Field(s)!")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox56.Clear()
        TextBox57.Clear()
    End Sub

    Private Sub Button8_Click_2(sender As Object, e As EventArgs) Handles Button8.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadPaymentMethods(con)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        customerListForm.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
        LoadEmployees(con)
    End Sub

    Function IsInDatagridview(ByVal cell1 As String, ByVal rowCell1_ID As Integer, ByRef dgv As DataGridView)

        Dim isFound As Boolean = False

        For Each rw As DataGridViewRow In dgv.Rows
            If rw.Cells(rowCell1_ID).Value.ToString = cell1 Then
                isFound = True
                Return True

            End If
        Next
        Return False
    End Function

    Private Sub ReturningCustomer(ByVal connectionString As String)
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim con As String = "Data Source=KEVINSPC;Initial Catalog=391back;Integrated Security=True;MultipleActiveResultSets=False"
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction
            transaction = connection.BeginTransaction("ReturningCustomer")
            command.Connection = connection
            command.Transaction = transaction
            Dim name As String = nameField.Text
            Dim cid As String = cidField.Text
            Dim country As String = countryField.Text
            Dim address As String = addressField.Text
            Dim city As String = cityField.Text
            Dim phone As String = phoneField.Text
            Dim post As String = postalCodeField.Text
            Dim prov As String = provinceField.Text
            Dim i = availableRoomTable.CurrentRow.Index
            Dim RID = availableRoomTable.Item(0, i).Value
            Dim dt1 As String = checkinDate.Value.Year & "-" & checkinDate.Value.Month & "-" & checkinDate.Value.Day
            Dim time1 As DateTime = DateTime.Parse(dt1)
            Dim dt2 As String = checkoutDate.Value.Year & "-" & checkoutDate.Value.Month & "-" & checkoutDate.Value.Day
            Dim time2 As DateTime = DateTime.Parse(dt2)
            Try


                command.CommandText = _
                  "Insert into Reservation (C_ID, RES_ID, R_ID, Check_In_Date, Check_Out_Date) VALUES (" & "'" & cid & "'," & "'" & reservationCountLabel.Text.ToString & "'," & "'" & RID & "'," & "'" & time1 & "'," & "'" & time2 & "')"
                command.ExecuteNonQuery()

                transaction.Commit()
                committedToDB()
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
            ReservationCount(con)
            PersonCount(con)
        End Using
    End Sub


    Private Sub Warehouse(ByVal connectionString As String, ByVal query As String)
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            adapter = New SqlDataAdapter(query, connection)
            adapter.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0)
            Dim command As SqlCommand = connection.CreateCommand()
            Dim transaction As SqlTransaction

            transaction = connection.BeginTransaction("Warehouse")

            command.Connection = connection
            command.Transaction = transaction

            Try
                command.CommandText = _
                    query
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

    Private Function buildQuery(ByVal str1 As String, ByVal str2 As String) As String
        If (str1 = "Service Department") Then
            If (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Month"
            ElseIf (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Quarter"
            ElseIf (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Year"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department, Employee.Province"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Services.Department"
            End If
        End If
        If (str1 = "Month") Then
            If (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Quarter"
            ElseIf (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Year"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Month"
            End If
        End If
        If (str1 = "Year") Then
            If (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Quarter"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Year"
            End If
        End If

        If (str1 = "Quarter") Then
            If (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Year"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Quarter"
            End If
        End If

        If (str1 = "Client Country") Then
            If (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Year"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Client_Origin.City"
            ElseIf (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By  Client_Origin.Country,Quarter"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Country"
            End If
        End If

        If (str1 = "Client City") Then
            If (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Year"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Month"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Client_Origin.Country"
            ElseIf (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By  Client_Origin.City,Quarter"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Client_Origin.Province"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.City"
            End If
        End If

        If (str1 = "Client Province") Then
            If (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Quarter"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Client_Origin.Country"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By  Client_Origin.Province, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Employee.Department"
            ElseIf (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Client_Origin.Province"
            End If
        End If

        If (str1 = "Employee Department") Then
            If (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Quarter"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Employee.Country"
            ElseIf (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Year"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Department"
            End If
        End If

        If (str1 = "Employee City") Then
            If (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Quarter"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Client_Origin.Province"
            ElseIf (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Year"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.City"
            End If
        End If

        If (str1 = "Employee Province") Then
            If (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Quarter"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Employee.City"
            ElseIf (str2 = "Employee Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Employee.Country"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Employee.Department"
            ElseIf (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Year"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Province"
            End If
        End If

        If (str1 = "Employee Country") Then
            If (str2 = "Quarter") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Quarter From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Quarter"
            ElseIf (str2 = "Month") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Month From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Month"
            ElseIf (str2 = "Client City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Client_Origin.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Client_Origin.City"
            ElseIf (str2 = "Client Country") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Client_Origin.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Client_Origin.Country"
            ElseIf (str2 = "Client Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Client_Origin.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Client_Origin.Province"
            ElseIf (str2 = "Employee City") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Employee.City From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Employee.City"
            ElseIf (str2 = "Year") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Year From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Year"
            ElseIf (str2 = "Employee Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Employee.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Employee.Department"
            ElseIf (str2 = "Employee Province") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Employee.Province From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Employee.Province"
            ElseIf (str2 = "Service Department") Then
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country, Services.Department From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country, Services.Department"
            Else
                buildQuery = "Select Sum(Total_Rental_Days) As 'Total Rental Days', Employee.Country From Rentals, Services, Date, Client_Origin, Employee Where Services.ID = Rentals.Services_ID and Rentals.Date_ID = Date.ID and Client_Origin.ID = Rentals.Client_ID and Employee.ID = Rentals.Employee_ID Group By Employee.Country"
            End If
        End If
    End Function

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim con As String = "Data Source=kevinspc;Initial Catalog=391starschema;Integrated Security=True"
        Warehouse(con, buildQuery(ComboBox4.Text, ComboBox5.Text))
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim con As String = "Data Source=kevinspc;Initial Catalog=391starschema;Integrated Security=True"
        Warehouse(con, buildQuery(ComboBox4.Text, ComboBox5.Text))
    End Sub

End Class


