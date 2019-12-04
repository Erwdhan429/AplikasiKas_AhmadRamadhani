Imports System.Data.Odbc

Public Class Form1
    Public tampil As Integer
    Public saldo As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TampilGrid()
        MunculCombo()
        updateDisplay()

    End Sub

    Private Sub TextBox1_KeyPress2(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 7
        If e.KeyChar = Chr(13) Then
            bukakoneksi()
            CMD = New OdbcCommand("Select * From transaksi where id ='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("id tidak ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("jumlah")
                TextBox3.Text = RD.Item("Keterangan")
                DTP_time.Value = RD.Item("tanggal")
                ComboBox1.Text = RD.Item("jenis")
                TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            bukakoneksi()
            If ComboBox1.Text = "Masuk" Then
                saldo = CInt(TextBox2.Text)
                tampil = tampil + saldo
            ElseIf ComboBox1.Text = "Keluar" Then
                saldo = CInt(TextBox2.Text)
                tampil = tampil - saldo
            End If

            If tampil < 0 Then
                tampil = tampil + saldo
                MsgBox("data minus")
            Else
                Dim simpan As String = "Insert into transaksi values('" & DTP_time.Value & "','" & TextBox1.Text & "','" & ComboBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')"

                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Input data berhasil")
            End If

            updateDisplay()
            TampilGrid()
            KosongkanData()
            tutupkoneksi()

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button2.Click

        bukakoneksi()

        Dim edit As String = "update transaksi set
        jumlah ='" & TextBox2.Text & "',
        keterangan ='" & TextBox3.Text & "',
        tanggal ='" & DTP_time.Value & "',
        jenis='" & ComboBox1.Text & "' where id='" & TextBox1.Text & "'"
        CMD = New OdbcCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Data berhasil di Update")

        TampilGrid()
        KosongkanData()
        tutupkoneksi()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih data yang akan dihapus dengan masukkan ID dan Enter")
        Else
            If MessageBox.Show("Yakin akan dihapus ?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                bukakoneksi()
                Dim hapus As String = "delete From transaksi where id='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                updateDisplay()
                TampilGrid()
                KosongkanData()
                tutupkoneksi()
            End If
        End If
    End Sub

    Sub TampilGrid()

        bukakoneksi()

        DA = New OdbcDataAdapter("select * From transaksi", CONN)
        DS = New DataSet
        DA.Fill(DS, "transaksi")
        DataGridView1.DataSource = DS.Tables("transaksi")

        tutupkoneksi()
    End Sub

    Sub MunculCombo()
        ComboBox1.Items.Add("Masuk")
        ComboBox1.Items.Add("Keluar")
    End Sub

    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub updateDisplay()
        bukakoneksi()
        CMD = New OdbcCommand(" select * from transaksi order by id desc limit 1", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            lbl_saldo.Text = "Rp. 0"
        Else
            lbl_saldo.Text = "RP. " & RD.Item("saldo")
            tampil = RD.Item("saldo")
        End If
        tutupkoneksi()
    End Sub
End Class