Imports System.Data.OleDb
Public Class Form2
    Dim cnnOLEDB As New OleDbConnection
    Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & System.Environment.CurrentDirectory & "\Database2.mdb"

    Private Sub btnSearch2_Click(sender As Object, e As EventArgs) Handles btnSearch2.Click
        cnnOLEDB.Open()
        Dim search As String
        search = "SELECT Entry, Name, Age, Date, Contact, Address, Temperature, Question1, Question2, Question3, Question4 FROM Data WHERE ENTRY ='" + txtSearch.Text + "'"
        Dim cmdOLEDB As New OleDbCommand(search, cnnOLEDB)
        Dim rdrOLEDB As OleDbDataReader
        rdrOLEDB = cmdOLEDB.ExecuteReader
        rdrOLEDB.Read()

        txtShowName.Text = rdrOLEDB("Name")
        txtShowAge.Text = rdrOLEDB("Age")
        txtShowDate.Text = rdrOLEDB("Date")
        txtShowContact.Text = rdrOLEDB("Contact")
        txtShowAddress.Text = rdrOLEDB("Address")
        txtShowTemperature.Text = rdrOLEDB("Temperature")
        TextBox1.Text = rdrOLEDB("Question1")
        TextBox2.Text = rdrOLEDB("Question2")
        TextBox3.Text = rdrOLEDB("Question3")
        TextBox4.Text = rdrOLEDB("Question4")

        cnnOLEDB.Close()

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cnnOLEDB.ConnectionString = strConnectionString
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Hide()
        Form1.Show()

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub
End Class