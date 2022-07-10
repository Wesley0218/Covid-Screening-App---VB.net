Imports System.Data.OleDb
Public Class Form1
    Dim cnnOLEDB As New OleDbConnection
    Dim Q1 As String
    Dim Q2 As String
    Dim Q3 As String
    Dim Q4 As String
    Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & System.Environment.CurrentDirectory & "\Database2.mdb"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cnnOLEDB.ConnectionString = strConnectionString


    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click

        If txtEntry.Text = Nothing Or txtName.Text = Nothing Or txtAge.Text = Nothing Or txtAddress.Text = Nothing Or txtDate.Text = Nothing Or txtContact.Text = Nothing Or txtTemperature.Text = Nothing Then
            MessageBox.Show("Fill-up form is not Complete", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If cnnOLEDB.State = ConnectionState.Closed Then
            cnnOLEDB.Open()

        End If


        Using create As New OleDbCommand("INSERT INTO Data([Entry],[Name], [Age], [Date], [Contact], [Address], [Temperature],[Question1],[Question2],[Question3],[Question4])VALUES(@ENTRY,@NAME,@AGE,@DATE,@CONTACT,@ADDRESS,@TEMPERATURE,@Q1,@Q2,@Q3,@Q4)", cnnOLEDB)
            create.Parameters.AddWithValue("@ENTRY", OleDbType.VarChar).Value = txtEntry.Text.Trim
            create.Parameters.AddWithValue("@NAME", OleDbType.VarChar).Value = txtName.Text.Trim
            create.Parameters.AddWithValue("@AGE", OleDbType.VarChar).Value = txtAge.Text.Trim
            create.Parameters.AddWithValue("@DATE", OleDbType.VarChar).Value = txtDate.Text.Trim
            create.Parameters.AddWithValue("@CONTACT", OleDbType.VarChar).Value = txtContact.Text.Trim
            create.Parameters.AddWithValue("@ADDRESS", OleDbType.VarChar).Value = txtAddress.Text.Trim
            create.Parameters.AddWithValue("@TEMPERATURE", OleDbType.VarChar).Value = txtTemperature.Text.Trim
            create.Parameters.AddWithValue("@Q1", OleDbType.VarChar).Value = Q1.Trim
            create.Parameters.AddWithValue("@Q2", OleDbType.VarChar).Value = Q2.Trim
            create.Parameters.AddWithValue("@Q3", OleDbType.VarChar).Value = Q3.Trim
            create.Parameters.AddWithValue("@Q4", OleDbType.VarChar).Value = Q4.Trim

            If create.ExecuteNonQuery Then
                MessageBox.Show("SUBMITTED SUCCESSFULLY!!!", "INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtName.Text = ""
                txtEntry.Text = ""
                txtAge.Text = ""
                txtDate.Text = ""
                txtContact.Text = ""
                txtAddress.Text = ""
                txtTemperature.Text = ""

            End If
        End Using
        Using create As New OleDbCommand("SELECT COUNT(*) FROM Data WHERE [ID] = @ID", cnnOLEDB)
            create.Parameters.AddWithValue("@ID", OleDbType.VarChar).Value = txtEntry.Text.Trim

            Dim createcount = Convert.ToInt32(create.ExecuteScalar())

            If createcount > 0 Then ' 
                MessageBox.Show("ENTRY NUMBER IS ALREADY EXIST!!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End Using

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        btnSearch.Enabled = False
        Form2.Show()


    End Sub

    Private Sub G1yes_CheckedChanged(sender As Object, e As EventArgs) Handles G1yes.CheckedChanged
        Q1 = "Yes"
    End Sub

    Private Sub G1no_CheckedChanged(sender As Object, e As EventArgs) Handles G1no.CheckedChanged
        Q1 = "No"
    End Sub

    Private Sub G2yes_CheckedChanged(sender As Object, e As EventArgs) Handles G2yes.CheckedChanged
        Q2 = "Yes"
    End Sub

    Private Sub G2no_CheckedChanged(sender As Object, e As EventArgs) Handles G2no.CheckedChanged
        Q2 = "No"
    End Sub

    Private Sub G3yes_CheckedChanged(sender As Object, e As EventArgs) Handles G3yes.CheckedChanged
        Q3 = "Yes"
    End Sub

    Private Sub G3no_CheckedChanged(sender As Object, e As EventArgs) Handles G3no.CheckedChanged
        Q3 = "No"
    End Sub

    Private Sub G4yes_CheckedChanged(sender As Object, e As EventArgs) Handles G4yes.CheckedChanged
        Q4 = "Yes"
    End Sub

    Private Sub G4no_CheckedChanged(sender As Object, e As EventArgs) Handles G4no.CheckedChanged
        Q4 = "No"
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter
        If G4no.Checked = True Then
            headache.Checked = False
            nose.Checked = False
            breath.Checked = False
            swallowing.Checked = False
            throat.Checked = False
            bodyaches.Checked = False
            cough.Checked = False
            smelltaste.Checked = False
            fatigue.Checked = False
            voice.Checked = False
        End If
    End Sub
End Class
