<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Room_ID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Floor = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Room_Type = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Num_Of_Beds = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.TextBox9 = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TextBox10 = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.TextBox11 = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.TextBox12 = New System.Windows.Forms.TextBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(468, 125)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(90, 13)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "Room Information"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Room_ID, Me.Floor, Me.Room_Type, Me.Num_Of_Beds})
        Me.DataGridView1.Location = New System.Drawing.Point(471, 151)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(452, 452)
        Me.DataGridView1.TabIndex = 0
        '
        'Room_ID
        '
        Me.Room_ID.HeaderText = "Room ID"
        Me.Room_ID.Name = "Room_ID"
        '
        'Floor
        '
        Me.Floor.HeaderText = "Floor"
        Me.Floor.Name = "Floor"
        '
        'Room_Type
        '
        Me.Room_Type.HeaderText = "Room Type"
        Me.Room_Type.Name = "Room_Type"
        '
        'Num_Of_Beds
        '
        Me.Num_Of_Beds.HeaderText = "Number Of Beds"
        Me.Num_Of_Beds.Name = "Num_Of_Beds"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label8.Location = New System.Drawing.Point(124, 10)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(130, 15)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Reservation Number: 001"
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(678, 609)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(75, 23)
        Me.Button8.TabIndex = 15
        Me.Button8.Text = "Clear"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(759, 609)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(75, 23)
        Me.Button7.TabIndex = 14
        Me.Button7.Text = "Cancel"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(840, 609)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(83, 23)
        Me.Button6.TabIndex = 13
        Me.Button6.Text = "Confirm"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(124, 153)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(245, 20)
        Me.TextBox6.TabIndex = 11
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(124, 129)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(245, 20)
        Me.TextBox5.TabIndex = 10
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(124, 105)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(245, 20)
        Me.TextBox4.TabIndex = 9
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(124, 83)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(245, 20)
        Me.TextBox3.TabIndex = 8
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(124, 60)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(245, 20)
        Me.TextBox2.TabIndex = 7
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(23, 125)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(106, 13)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "Customer Information"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(124, 37)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(245, 20)
        Me.TextBox1.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(72, 86)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 13)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Country:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(50, 63)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Customer ID:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(70, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Address:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(80, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Name:"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TextBox12)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.TextBox11)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.TextBox10)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.TextBox9)
        Me.Panel1.Controls.Add(Me.TextBox8)
        Me.Panel1.Controls.Add(Me.TextBox7)
        Me.Panel1.Controls.Add(Me.DateTimePicker2)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.TextBox6)
        Me.Panel1.Controls.Add(Me.TextBox5)
        Me.Panel1.Controls.Add(Me.TextBox4)
        Me.Panel1.Controls.Add(Me.TextBox3)
        Me.Panel1.Controls.Add(Me.TextBox2)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(23, 151)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(442, 452)
        Me.Panel1.TabIndex = 18
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(37, 156)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 13)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Phone Number:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(91, 132)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(27, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "City:"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(848, 37)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 23)
        Me.Button5.TabIndex = 17
        Me.Button5.Text = "Sign Out"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(398, 70)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(115, 23)
        Me.Button4.TabIndex = 16
        Me.Button4.Text = "View Inventory"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(252, 70)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(140, 23)
        Me.Button3.TabIndex = 15
        Me.Button3.Text = "View All Customers"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(147, 70)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(99, 23)
        Me.Button2.TabIndex = 14
        Me.Button2.Text = "Checkout"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(23, 70)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(118, 23)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "Book a Customer"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "HOTEL_NAME"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(696, 42)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(146, 13)
        Me.Label11.TabIndex = 22
        Me.Label11.Text = "Logged In As: Employee_000"
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(527, 70)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(120, 23)
        Me.Button9.TabIndex = 23
        Me.Button9.Text = "View Employees"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(36, 231)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(79, 13)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "Check In Time:"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(28, 253)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(87, 13)
        Me.Label13.TabIndex = 18
        Me.Label13.Text = "Check Out Time:"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(40, 183)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(78, 13)
        Me.Label14.TabIndex = 19
        Me.Label14.Text = "Check in Date:"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(32, 276)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(83, 13)
        Me.Label15.TabIndex = 20
        Me.Label15.Text = "Selected Room:"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Location = New System.Drawing.Point(124, 177)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(245, 20)
        Me.DateTimePicker1.TabIndex = 22
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(36, 208)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(82, 13)
        Me.Label16.TabIndex = 23
        Me.Label16.Text = "Checkout Date:"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Location = New System.Drawing.Point(124, 202)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(245, 20)
        Me.DateTimePicker2.TabIndex = 24
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(124, 224)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(245, 20)
        Me.TextBox7.TabIndex = 25
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(124, 246)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(245, 20)
        Me.TextBox8.TabIndex = 26
        '
        'TextBox9
        '
        Me.TextBox9.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox9.Location = New System.Drawing.Point(124, 269)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(66, 20)
        Me.TextBox9.TabIndex = 27
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(207, 276)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(94, 13)
        Me.Label17.TabIndex = 28
        Me.Label17.Text = "No. of Occupants:"
        '
        'TextBox10
        '
        Me.TextBox10.Location = New System.Drawing.Point(307, 269)
        Me.TextBox10.Name = "TextBox10"
        Me.TextBox10.Size = New System.Drawing.Size(62, 20)
        Me.TextBox10.TabIndex = 29
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(80, 298)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(34, 13)
        Me.Label18.TabIndex = 30
        Me.Label18.Text = "Type:"
        '
        'TextBox11
        '
        Me.TextBox11.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox11.Location = New System.Drawing.Point(124, 295)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.Size = New System.Drawing.Size(245, 20)
        Me.TextBox11.TabIndex = 31
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(9, 398)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(109, 13)
        Me.Label19.TabIndex = 32
        Me.Label19.Text = "Description/Features:"
        '
        'TextBox12
        '
        Me.TextBox12.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox12.Location = New System.Drawing.Point(124, 321)
        Me.TextBox12.Multiline = True
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.Size = New System.Drawing.Size(245, 90)
        Me.TextBox12.TabIndex = 33
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(961, 665)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Room_ID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Floor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Room_Type As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Num_Of_Beds As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label

End Class
