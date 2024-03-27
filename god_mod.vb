Imports MySql.Data.MySqlClient

Module god_mod
    Dim con As New MySqlConnection("server=127.0.0.1;uid=root;database=mysql;")
    Dim con2 As New MySqlConnection("server=127.0.0.1;uid=root;database=mysql;")
    Dim com As New MySqlCommand
    Dim dr As MySqlDataReader

    Sub Opencon()
        con.Close()
        con.Open()
        com.Connection = con
        con2.Close()
        con2.Open()
        com.Connection = con2
    End Sub

    Sub CloseReader()
        dr = com.ExecuteReader
        dr.Close()
    End Sub

    Sub Create_database_bless()
        com = New MySqlCommand("create database if not exists db_bless;", con)
        CloseReader()
    End Sub

    Sub Create_database_bless2()
        com = New MySqlCommand("create database if not exists db_bless;", con2)
        CloseReader()
    End Sub

    Sub Create_table_donation()
        com = New MySqlCommand("create table if not exists db_bless.donors(id int(100) auto_increment Primary key, date text, time text, firstname text, middlename text, lastname text, age int(2), address text, sex text, quantity int(2), bloodtype text, addedby text, updatedby text, dateofupdate text)", con)
        CloseReader()
    End Sub

    Sub Deleteall()
        com = New MySqlCommand("drop table db_bless.donors;", con)
        CloseReader()
    End Sub

    Sub Donation_params()
        com.Parameters.AddWithValue("@date", Form1.ToolStripStatusLabel3.Text)
        com.Parameters.AddWithValue("@time", Form1.ToolStripStatusLabel4.Text)
        com.Parameters.AddWithValue("@firstname", Form1.TextBox1.Text)
        com.Parameters.AddWithValue("@middlename", Form1.TextBox2.Text)
        com.Parameters.AddWithValue("@lastname", Form1.TextBox3.Text)
        com.Parameters.AddWithValue("@age", Form1.TextBox4.Text)
        com.Parameters.AddWithValue("@address", Form1.TextBox5.Text)
        com.Parameters.AddWithValue("@sex", Form1.ComboBox1.Text)
        com.Parameters.AddWithValue("@quantity", Form1.TextBox6.Text)
        com.Parameters.AddWithValue("@bloodtype", Form1.ComboBox2.Text)
        com.Parameters.AddWithValue("@addedby", Form1.ToolStripStatusLabel6.Text)
        com.Parameters.AddWithValue("@updatedby", Form1.ToolStripStatusLabel6.Text)
        com.Parameters.AddWithValue("@dateofupdate", Form1.ToolStripStatusLabel3.Text & "|" & Form1.ToolStripStatusLabel4.Text)
    End Sub


    Sub Add_new_donor()
        com = New MySqlCommand("insert into db_bless.donors(date, time, firstname, middlename, lastname, age, address, sex, quantity, bloodtype, addedby)values(@date,  @time, @firstname, @middlename, @lastname, @age, @address, @sex, @quantity, @bloodtype, @addedby)", con)
        Donation_params()
        CloseReader()
    End Sub


    Sub Donors_subs()
        Dim dn As New ListViewItem(dr.Item(0).ToString)
        Dim dnc As Integer
        For dnc = 1 To 11 Step 1
            dn.SubItems.Add(dr.Item(dnc).ToString).BackColor = Color.White
        Next
        dn.UseItemStyleForSubItems = False
        Form1.ListView1.Items.Add(dn)
    End Sub

    Sub Auto_counts()
        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""A+"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label3.Text = dr.Item(0).ToString
        End While
        dr.Close()

        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""AB+"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label4.Text = dr.Item(0).ToString
        End While
        dr.Close()

        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""B+"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label5.Text = dr.Item(0).ToString
        End While
        dr.Close()

        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""O+"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label6.Text = dr.Item(0).ToString
        End While
        dr.Close()


        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""A-"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label7.Text = dr.Item(0).ToString
        End While
        dr.Close()

        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""AB-"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label8.Text = dr.Item(0).ToString
        End While
        dr.Close()

        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""B-"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label9.Text = dr.Item(0).ToString
        End While
        dr.Close()

        com = New MySqlCommand("select sum(quantity) from db_bless.donors where bloodtype=""O-"";", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label10.Text = dr.Item(0).ToString
        End While
        dr.Close()


        com = New MySqlCommand("select sum(quantity) from db_bless.donors;", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.Label19.Text = dr.Item(0).ToString
        End While
        dr.Close()

    End Sub

    Sub Reload_donors()
        Form1.ListView1.Items.Clear()
        com = New MySqlCommand("select * from db_bless.donors order by id desc;", con)
        dr = com.ExecuteReader
        While dr.Read
            Donors_subs()
        End While
        Form1.Label11.Text = "Total No. of Donors: " & Form1.ListView1.Items.Count
        Form1.Label24.Text = Form1.ListView1.Items.Count
        dr.Close()
        Auto_counts()
    End Sub

    Sub Select_donor()
        com = New MySqlCommand("select * from db_bless.donors  where id='" & Form1.ListView1.FocusedItem.Text & "';", con)
        dr = com.ExecuteReader
        While dr.Read
            Form1.TextBox1.Text = dr.Item(3).ToString
            Form1.TextBox2.Text = dr.Item(4).ToString
            Form1.TextBox3.Text = dr.Item(5).ToString
            Form1.TextBox4.Text = dr.Item(6).ToString
            Form1.TextBox5.Text = dr.Item(7).ToString
            Form1.ComboBox1.Text = dr.Item(8).ToString
            Form1.TextBox6.Text = dr.Item(9).ToString
            Form1.ComboBox2.Text = dr.Item(10).ToString
        End While
        dr.Close()
    End Sub

    Sub Update_donor()
        com = New MySqlCommand("update db_bless.donors set id='" & Form1.ListView1.FocusedItem.Text & "', firstname=@firstname, middlename=@middlename, lastname=@lastname, age=@age, address=@address, sex=@sex, quantity=@quantity, bloodtype=@bloodtype where id='" & Form1.ListView1.FocusedItem.Text & "';", con)
        Donation_params()
        CloseReader()
    End Sub

    Sub Delete_donor()
        com = New MySqlCommand("delete from db_bless.donors where id='" & Form1.ListView1.FocusedItem.Text & "';", con)
        CloseReader()
    End Sub

    Sub Search_donors()
        Form1.ListView1.Items.Clear()
        com = New MySqlCommand("select * from db_bless.donors where firstname like '%" & Form1.TextBox13.Text & "%' or lastname like '%" & Form1.TextBox13.Text & "%' or bloodtype = '" & Form1.TextBox13.Text & "' order by id desc;", con)
        dr = com.ExecuteReader
        While dr.Read
            Donors_subs()
        End While
        Form1.Label11.Text = "Total No. of Donors: " & Form1.ListView1.Items.Count
        Form1.Label24.Text = Form1.ListView1.Items.Count
        dr.Close()
    End Sub

    Sub Search_donors_bydate()
        Form1.ListView1.Items.Clear()
        com = New MySqlCommand("select * from db_bless.donors where date='" & Form1.DateTimePicker1.Text & "' order by id desc;", con)
        dr = com.ExecuteReader
        While dr.Read
            Donors_subs()
        End While
        Form1.Label11.Text = "Total No. of Donors: " & Form1.ListView1.Items.Count
        Form1.Label24.Text = Form1.ListView1.Items.Count
        dr.Close()
    End Sub

    Sub Filter_donors_bydate()
        Form1.ListView1.Items.Clear()
        com = New MySqlCommand("select * from db_bless.donors where date between '" & Form1.DateTimePicker1.Text & "' and '" & Form1.DateTimePicker2.Text & "' order by id desc;", con)
        dr = com.ExecuteReader
        While dr.Read
            Donors_subs()
        End While
        Form1.Label11.Text = "Total No. of Donors: " & Form1.ListView1.Items.Count
        Form1.Label24.Text = Form1.ListView1.Items.Count
        dr.Close()
    End Sub











End Module
