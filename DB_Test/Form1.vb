Imports System.Data.SQLite
Imports System.Net.Sockets
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.IO.Ports


Public Class Form1



    Public dbfile As String
    Public drink(5000) As String
    Public drink_id(5000) As Integer
    Public recipe(6000, 30) As Single
    Public active_recipe(15) As Single
    Public ing_names(500) As String
    Public ing_data(500, 3) As Single
    Public comments(6000) As String
    Public idx As Integer = 0
    Public active_ing(8) As Integer
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Public scale_size As Single = 1
    Public total As Single
    Public SER_PRT As String
    Public cmd_string_long As String



    Private Sub lst_records_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lst_records.SelectedIndexChanged

        Dim curItem_s As String = Microsoft.VisualBasic.Left(lst_records.SelectedItem, 4)
        Dim curitem As Integer = Val(curItem_s)


        lst_ing.Items.Clear()
        lst_amt.Items.Clear()

        ' Populate Ingredients, Totalize mix amount
        total = 0

        For c = 0 To 29 Step 2
            If recipe(curitem, c + 1) = 0 Then Exit For
            lst_ing.Items.Add(String.Format("{0}", ing_names(recipe(curitem, c))))
            total = total + recipe(curitem, c + 1)

        Next

        'populate list with percentages

        For c = 0 To 29 Step 2

            If recipe(curitem, c + 1) = 0 Then Exit For
            lst_amt.Items.Add(String.Format("{0:P}", ((recipe(curitem, c + 1) / total))))

        Next


        txt_comments.Clear()
        txt_comments.Text = comments(curitem)
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Load_DB()
    End Sub

    Private Sub Load_DB()

        Dim f As New OpenFileDialog
        f.Filter = "SQLite 3 (*.db3)|*.db3|All Files|*.*"
        If f.ShowDialog() = DialogResult.OK Then
            Dim SQLconnect As New SQLite.SQLiteConnection()
            Dim SQLcommand As SQLiteCommand
            SQLconnect.ConnectionString = "Data Source=" & f.FileName & ";"
            dbfile = f.FileName
            SQLconnect.Open()

            ' Populate recipe array
            SQLcommand = SQLconnect.CreateCommand
            SQLcommand.CommandText = "select drink_id, position, ing_id, amount from recipes order by drink_id"
            Dim SQLreader2 As SQLiteDataReader = SQLcommand.ExecuteReader()

            While SQLreader2.Read()
                recipe(SQLreader2("drink_id"), SQLreader2("Position") * 2) = SQLreader2("Ing_ID")
                recipe(SQLreader2("drink_id"), (SQLreader2("Position") * 2) + 1) = SQLreader2("Amount")
            End While

            ' Populate ingredient array
            SQLcommand = SQLconnect.CreateCommand
            SQLcommand.CommandText = "select id, name, bottleposition from ingredients"
            Dim SQLreader3 As SQLiteDataReader = SQLcommand.ExecuteReader()

            While SQLreader3.Read()
                ing_names(SQLreader3("id")) = SQLreader3("Name")
            End While

            SQLcommand = SQLconnect.CreateCommand
            SQLcommand.CommandText = "SELECT * FROM DrinkNames ORDER BY Name ASC"

            Dim SQLreader As SQLiteDataReader = SQLcommand.ExecuteReader()
            While SQLreader.Read()
                drink(idx) = SQLreader("Name")
                drink_id(idx) = SQLreader("ID")
                comments(drink_id(idx)) = SQLreader("Comments")
                idx = idx + 1
            End While
            ' MessageBox.Show(idx)
            SQLcommand.Dispose()
            SQLconnect.Close()

            'populate ingredients listboxes
            Call pop_ing_listboxes()

            'populate drinks listbox
            Call pop_drinks_listbox()


        End If
    End Sub

    Private Sub pop_drinks_listbox()
        Dim valid_ing(29) As Boolean
        lst_records.Items.Clear()

        For c = 0 To idx - 1

            For c1 = 0 To 29 Step 2
                If ((recipe(drink_id(c), c1) = (active_ing(1)) Or (recipe(drink_id(c), c1) = active_ing(2)) Or (recipe(drink_id(c), c1) = active_ing(3)) Or (recipe(drink_id(c), c1) = active_ing(4)) Or (recipe(drink_id(c), c1) = active_ing(5)) Or (recipe(drink_id(c), c1) = active_ing(6)) Or (recipe(drink_id(c), c1) = active_ing(7)) Or (recipe(drink_id(c), c1) = active_ing(8)) Or (recipe(drink_id(c), c1) = 0))) Then valid_ing(c1) = True
            Next c1

            If valid_ing(0) And valid_ing(2) And valid_ing(4) And valid_ing(6) And valid_ing(8) And valid_ing(10) And valid_ing(12) And valid_ing(14) And valid_ing(16) And valid_ing(18) And valid_ing(20) And valid_ing(22) And valid_ing(24) And valid_ing(26) And valid_ing(28) Then
                lst_records.Items.Add(String.Format("{1,-10}  {0}", drink(c), drink_id(c)))
            End If

            For c1 = 0 To 29 Step 2
                valid_ing(c1) = False
            Next c1

        Next c
    End Sub

    Private Sub pop_ing_listboxes()
        cmb_ing1.Items.Clear()
        For c = 1 To 500
            cmb_ing1.Items.Add(String.Format("{0}", ing_names(c)))
            cmb_ing2.Items.Add(String.Format("{0}", ing_names(c)))
            cmb_ing3.Items.Add(String.Format("{0}", ing_names(c)))
            cmb_ing4.Items.Add(String.Format("{0}", ing_names(c)))
            cmb_ing5.Items.Add(String.Format("{0}", ing_names(c)))
            cmb_ing6.Items.Add(String.Format("{0}", ing_names(c)))
            cmb_ing7.Items.Add(String.Format("{0}", ing_names(c)))
            cmb_ing8.Items.Add(String.Format("{0}", ing_names(c)))


        Next c

        active_ing(1) = 81
        active_ing(2) = 31
        active_ing(3) = 139
        active_ing(4) = 32
        active_ing(5) = 39
        active_ing(6) = 5
        active_ing(7) = 21
        active_ing(8) = 25

        cmb_ing1.Text = ing_names(active_ing(1))
        cmb_ing2.Text = ing_names(active_ing(2))
        cmb_ing3.Text = ing_names(active_ing(3))
        cmb_ing4.Text = ing_names(active_ing(4))
        cmb_ing5.Text = ing_names(active_ing(5))
        cmb_ing6.Text = ing_names(active_ing(6))
        cmb_ing7.Text = ing_names(active_ing(7))
        cmb_ing8.Text = ing_names(active_ing(8))

    End Sub


    Private Sub cmb_ing1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing1.SelectedIndexChanged
        active_ing(1) = cmb_ing1.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub cmb_ing2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing2.SelectedIndexChanged
        active_ing(2) = cmb_ing2.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub cmb_ing3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing3.SelectedIndexChanged
        active_ing(3) = cmb_ing3.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub cmb_ing4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing4.SelectedIndexChanged
        active_ing(4) = cmb_ing4.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub cmb_ing5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing5.SelectedIndexChanged
        active_ing(5) = cmb_ing5.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub cmb_ing6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing6.SelectedIndexChanged
        active_ing(6) = cmb_ing6.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub cmb_ing7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing7.SelectedIndexChanged
        active_ing(7) = cmb_ing7.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub cmb_ing8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_ing8.SelectedIndexChanged
        active_ing(8) = cmb_ing8.SelectedIndex + 1
        Call pop_drinks_listbox()
    End Sub

    Private Sub but_makeit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_makeit.Click

        Dim curItem_s As String = Microsoft.VisualBasic.Left(lst_records.SelectedItem, 4)
        Dim curitem As Integer = Val(curItem_s)
        Dim field_counter As Integer
        cmd_string_long = ""
        scale_size = Math.Abs(Size_0.Checked * num_scale_0.Value + Size_1.Checked * num_scale_1.Value + Size_2.Checked * num_Scale_2.Value + Size_3.Checked * num_scale_3.Value)

        For c = 0 To 29 Step 2

            ' If (recipe(curitem, c) = active_ing(1)) Then Call send_bit(0, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(0, 0)
            ' If (recipe(curitem, c) = active_ing(2)) Then Call send_bit(1, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(1, 0)
            ' If (recipe(curitem, c) = active_ing(3)) Then Call send_bit(2, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(2, 0)
            ' If (recipe(curitem, c) = active_ing(4)) Then Call send_bit(3, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(3, 0)
            ' If (recipe(curitem, c) = active_ing(5)) Then Call send_bit(4, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(4, 0)
            ' If (recipe(curitem, c) = active_ing(6)) Then Call send_bit(5, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(5, 0)
            ' If (recipe(curitem, c) = active_ing(7)) Then Call send_bit(6, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(6, 0)
            ' If (recipe(curitem, c) = active_ing(8)) Then Call send_bit(7, ((recipe(curitem, c + 1) / total) * 100 * scale_size)) Else send_bit(7, 0)
            If (recipe(curitem, c) = active_ing(1)) Then cmd_string_long = cmd_string_long + "0," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1
            If (recipe(curitem, c) = active_ing(2)) Then cmd_string_long = cmd_string_long + "1," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1
            If (recipe(curitem, c) = active_ing(3)) Then cmd_string_long = cmd_string_long + "2," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1
            If (recipe(curitem, c) = active_ing(4)) Then cmd_string_long = cmd_string_long + "3," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1
            If (recipe(curitem, c) = active_ing(5)) Then cmd_string_long = cmd_string_long + "4," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1
            If (recipe(curitem, c) = active_ing(6)) Then cmd_string_long = cmd_string_long + "5," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1
            If (recipe(curitem, c) = active_ing(7)) Then cmd_string_long = cmd_string_long + "6," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1
            If (recipe(curitem, c) = active_ing(8)) Then cmd_string_long = cmd_string_long + "7," + Str(CUInt(((recipe(curitem, c + 1) / total) * 100 * scale_size))) : field_counter = field_counter + 1

            If ((recipe(curitem, c + 2)) <> 0) Then cmd_string_long = cmd_string_long + ","





        Next c
        For c = field_counter To 7
            cmd_string_long = cmd_string_long + ",0,0"
        Next

        'MsgBox(cmd_string_long)
        send_bit(0, 0)
        SerialPort1.Close()
        My.Computer.Audio.Play(My.Resources.buff_drunk, AudioPlayMode.Background)

    End Sub


    Private Sub send_bit(ByVal bit As Integer, ByVal time_delay As Long)
        Dim response As String
        On Error GoTo Errorhandler

        System.Windows.Forms.Application.DoEvents()

        cmd_string_long = cmd_string_long + vbNewLine
        SerialPort1.Open()
        SerialPort1.Write(cmd_string_long)
        cmd_string_long = ""

Wait_for_done:
        response = SerialPort1.ReadLine()
        If response = "DONE" Then SerialPort1.Close() : Exit Sub
        Me.Refresh()
        GoTo Wait_for_done
       

        Exit Sub

Errorhandler:
        MsgBox("Serial port not defined.")




    End Sub



    Private Sub QuitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuitToolStripMenuItem.Click
        End
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem.Click
        cmd_string_long = "0,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem1.Click
        cmd_string_long = "1,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem2.Click
        cmd_string_long = "2,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem3.Click
        cmd_string_long = "3,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem4.Click
        cmd_string_long = "4,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem5.Click
        cmd_string_long = "5,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem6.Click
        cmd_string_long = "6,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub PurgeIng1ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurgeIng1ToolStripMenuItem7.Click
        cmd_string_long = "7,1000,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        Call send_bit(0, 0)
    End Sub

    Private Sub LoadNewDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_DB()
    End Sub


    Private Sub COM1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles COM1ToolStripMenuItem.Click
        SER_PRT = "COM1"
        SerialPort1.Close()
        SerialPort1.PortName = SER_PRT 'change com port to match your Arduino port
        SerialPort1.BaudRate = 9600
        SerialPort1.DataBits = 8
        SerialPort1.Parity = Parity.None
        SerialPort1.StopBits = StopBits.One
        SerialPort1.Handshake = Handshake.None
        SerialPort1.Encoding = System.Text.Encoding.Default 'very important!
    End Sub

    Private Sub COM2ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles COM2ToolStripMenuItem.Click
        SER_PRT = "COM2"
        SerialPort1.Close()
        SerialPort1.PortName = SER_PRT 'change com port to match your Arduino port
        SerialPort1.BaudRate = 9600
        SerialPort1.DataBits = 8
        SerialPort1.Parity = Parity.None
        SerialPort1.StopBits = StopBits.One
        SerialPort1.Handshake = Handshake.None
        SerialPort1.Encoding = System.Text.Encoding.Default 'very important!
    End Sub

    Private Sub COM20ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles COM20ToolStripMenuItem.Click
        SER_PRT = "COM20"
        SerialPort1.Close()
        SerialPort1.PortName = SER_PRT 'change com port to match your Arduino port
        SerialPort1.BaudRate = 1200
        SerialPort1.DataBits = 8
        SerialPort1.Parity = Parity.None
        SerialPort1.StopBits = StopBits.One
        SerialPort1.Handshake = Handshake.None
        SerialPort1.Encoding = System.Text.Encoding.Default 'very important!
    End Sub

 

End Class