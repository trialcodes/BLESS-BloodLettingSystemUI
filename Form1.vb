
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        god_mod.Opencon()
        god_mod.Create_database_bless()
        god_mod.Create_table_donation()
        Timer1.Start()
        Cboxes()
        Lvs()
        Reload_donors()
        Auto_counts()
        Apn_perc()
        ABpn_perc()
        Bpn_perc()
        Opn_perc()
        Label68.Text = Label19.Text & "=100%"
        MetroTabPage1.Visible = True
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub

    Sub Apn_perc()
        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label3.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label33.Text = n3


            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label33.Width = Label70.Text


            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label33.Text
            n6 = n5 * 100
            Label33.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try

        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label7.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label32.Text = n3


            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label32.Width = Label70.Text


            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label32.Text
            n6 = n5 * 100
            Label32.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try

    End Sub

    Sub Opn_perc()
        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label6.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label27.Text = n3

            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label27.Width = Label70.Text

            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label27.Text
            n6 = n5 * 100
            Label27.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try

        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label10.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label26.Text = n3

            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label26.Width = Label70.Text

            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label26.Text
            n6 = n5 * 100
            Label26.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try
    End Sub

    Sub ABpn_perc()
        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label4.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label31.Text = n3

            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label31.Width = Label70.Text

            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label31.Text
            n6 = n5 * 100
            Label31.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try

        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label8.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label30.Text = n3

            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label30.Width = Label70.Text

            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label30.Text
            n6 = n5 * 100
            Label30.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try
    End Sub

    Sub Bpn_perc()
        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label5.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label29.Text = n3

            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label29.Width = Label70.Text

            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label29.Text
            n6 = n5 * 100
            Label29.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try

        Try
            Dim n1, n2, n3, n4 As Double
            n1 = Label9.Text
            n2 = Label19.Text
            n3 = n1 / n2
            Label28.Text = n3

            n4 = Label70.Text
            n4 = n4 * n3
            Label70.Text = n4
            Label28.Width = Label70.Text

            Dim n5 As Double
            Dim n6 As Integer
            n5 = Label28.Text
            n6 = n5 * 100
            Label28.Text = n6 & "%"
            Label70.Text = "300"
        Catch ex As Exception
        End Try
    End Sub

    Sub Lvs()
        With ListView1
            .View = View.Details
            .GridLines = False
            .BackColor = Color.DarkRed
            .FullRowSelect = True
            .HeaderStyle = ColumnHeaderStyle.Nonclickable
            .BorderStyle = BorderStyle.Fixed3D
            .Columns.Clear()
            .Columns.Add("#", 0)
            .Columns.Add("Date", 75)
            .Columns.Add("Time", 80)
            .Columns.Add("Firstname", 100)
            .Columns.Add("Middlename", 100)
            .Columns.Add("Lastname", 100)
            .Columns.Add("Age", 100)
            .Columns.Add("Address", 100)
            .Columns.Add("Sex", 100)
            .Columns.Add("Quantity", 100)
            .Columns.Add("Bloodtype", 100)
            .Columns.Add("Added by", 100)
        End With

        With ListView2
            .View = View.Details
            .GridLines = False

            .BackColor = Color.DarkRed
            .FullRowSelect = True
            .HeaderStyle = ColumnHeaderStyle.Nonclickable
            .BorderStyle = BorderStyle.Fixed3D
            .Columns.Clear()
            .Columns.Add("#", 0)
            .Columns.Add("Date", 100)
            .Columns.Add("Firstname", 100)
            .Columns.Add("Lastname", 100)
            .Columns.Add("Middlename", 100)
            .Columns.Add("Age", 100)
            .Columns.Add("Address", 100)
            .Columns.Add("Sex", 100)
            .Columns.Add("Quantity", 100)
            .Columns.Add("Bloodtype", 100)
            .Columns.Add("Added by", 100)
        End With
    End Sub

    Sub Cboxes()
        With ComboBox2
            .DropDownStyle = ComboBoxStyle.DropDownList
            .FlatStyle = FlatStyle.Popup
            .Items.Clear()
            .Items.Add("[ SELECT ]")
            .Items.Add("A+")
            .Items.Add("A-")
            .Items.Add("A+")
            .Items.Add("B+")
            .Items.Add("B-")
            .Items.Add("AB+")
            .Items.Add("AB-")
            .Items.Add("O+")
            .Items.Add("O-")
            .Text = "[ SELECT ]"
        End With

        With ComboBox3
            .DropDownStyle = ComboBoxStyle.DropDownList
            .FlatStyle = FlatStyle.Popup
            .Items.Clear()
            .Items.Clear()
            .Items.Add("[ SELECT ]")
            .Items.Add("A+")
            .Items.Add("A-")
            .Items.Add("A+")
            .Items.Add("B+")
            .Items.Add("B-")
            .Items.Add("AB+")
            .Items.Add("AB-")
            .Items.Add("O+")
            .Items.Add("O-")
            .Text = "[ SELECT ]"
        End With

        With ComboBox1
            .DropDownStyle = ComboBoxStyle.DropDownList
            .FlatStyle = FlatStyle.Popup
            .Items.Clear()
            .Items.Add("[ SELECT ]")
            .Items.Add("MALE")
            .Items.Add("FEMALE")
            .Text = "[ SELECT ]"
        End With

        With ComboBox4
            .DropDownStyle = ComboBoxStyle.DropDownList
            .FlatStyle = FlatStyle.Popup
            .Items.Clear()
            .Items.Add("[ SELECT ]")
            .Items.Add("MALE")
            .Items.Add("FEMALE")
            .Text = "[ SELECT ]"
        End With
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)
        If Not IsNumeric(TextBox4.Text) Then
            TextBox4.Clear()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel3.Text = Today
        ToolStripStatusLabel4.Text = TimeOfDay
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or ComboBox1.Text = "[ SELECT ]" Or ComboBox2.Text = "[ SELECT ]" Then
            MessageBox.Show("Please input all required information.", "Missing Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else



            Dim ask As MsgBoxResult = MessageBox.Show("Are you sure to delete that donor?", "Deleting", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If ask = MsgBoxResult.Yes Then
                Delete_donor()
                Reload_donors()
                MessageBox.Show("A donor deleted.", "Removed Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Txtclear()
            Else
                MessageBox.Show("Deleting a donor is now cancelled.", "Cancelled Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Txtclear()
            End If

        End If
        Apn_perc()
        ABpn_perc()
        Bpn_perc()
        Opn_perc()
        Label68.Text = Label19.Text & "=100%"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or ComboBox1.Text = "[ SELECT ]" Or ComboBox2.Text = "[ SELECT ]" Then
            MessageBox.Show("Please input all required information.", "Missing Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            Update_donor()
            Reload_donors()
            MessageBox.Show("A donor details now updated.", "Update Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Txtclear()
        End If
        Apn_perc()
        ABpn_perc()
        Bpn_perc()
        Opn_perc()
        Label68.Text = Label19.Text & "=100%"

    End Sub

    Sub Txtclear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        ComboBox1.Text = "[ SELECT ]"
        ComboBox2.Text = "[ SELECT ]"
        Reload_donors()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or ComboBox1.Text = "[ SELECT ]" Or ComboBox2.Text = "[ SELECT ]" Then
            MessageBox.Show("Please input all required information.", "Missing Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            Add_new_donor()
            Reload_donors()
            MessageBox.Show("A new donor added.", "Saved Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Txtclear()

        End If
        Apn_perc()
        ABpn_perc()
        Bpn_perc()
        Opn_perc()
        Label68.Text = Label19.Text & "=100%"
    End Sub

    Private Sub MetroTabPage1_Click(sender As Object, e As EventArgs) Handles MetroTabPage1.Click
        Apn_perc()
        ABpn_perc()
        Bpn_perc()
        Opn_perc()
        Label68.Text = Label19.Text & "=100%"
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        Select_donor()
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Txtclear()
        Reload_donors()
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        If TextBox13.Text = "" Then
            Reload_donors()
        ElseIf TextBox13.Text = "reset" Then
            Deleteall()
            MessageBox.Show("Table Donors Reset", "Reset Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Search_donors()
        End If

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Search_donors_bydate()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Try
            If DateTimePicker2.Value < DateTimePicker1.Value Then
                DateTimePicker2.Value = DateTimePicker1.Value
            End If
        Catch ex As Exception
        End Try
        Filter_donors_bydate()
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If TextBox9.Text = "" Then
        ElseIf Not IsNumeric(TextBox9.Text) Then
            TextBox9.Clear()
            MessageBox.Show("Only Numbers allowed for age.")
        End If
    End Sub

    Private Sub TextBox4_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox4.MaxLength = 2
        If TextBox4.Text = "" Then
        ElseIf Not IsNumeric(TextBox4.Text) Then
            TextBox4.Clear()
            MessageBox.Show("Only Numbers allowed for age.")
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        TextBox6.MaxLength = 1
        If TextBox6.Text = "" Then
        ElseIf Not IsNumeric(TextBox6.Text) Then
            TextBox6.Clear()
            MessageBox.Show("Only Numbers allowed for quantity/bags.")
        End If

    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click

    End Sub
End Class
