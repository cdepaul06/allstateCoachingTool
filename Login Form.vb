Imports System.Data.OleDb

Public Class frmAdmin
	Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
		'Verify that the user has entered credentials and if no credentials have been entered, display a messagebox.
		If txtUsername.Text Is "" Or txtPassword.Text Is "" Then
			MessageBox.Show("Please enter credentials before proceeding", "Admin Login", MessageBoxButtons.OK, MessageBoxIcon.Error)
			txtUsername.Focus()
			Exit Sub
		End If

		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Dim strUser, strPassword, strU, strP As String
		strUser = Me.txtUsername.Text
		strPassword = Me.txtPassword.Text

		Try
			conn.Open()
			Dim sql As String = "SELECT * FROM Admin WHERE username = '" & strUser & "' and password = '" & strPassword & "'"
			Dim cmd As New OleDbCommand(sql, conn)

			Dim sqlReader As OleDbDataReader = cmd.ExecuteReader
			Do While sqlReader.Read
				strU = sqlReader("username").ToString
				strP = sqlReader("password").ToString

				If strUser = strU And strPassword = strP Then
					'Hide form on successful login, open the splash/switchboard form and then close the login form.
					Me.Hide()
					frmUpdate.Show()
					Me.Close()
					Exit Sub
				End If
			Loop
			'If the username and password provided is not an authorized user show a messagebox.
			MessageBox.Show("Incorrect username or password.", "Admin Login", MessageBoxButtons.OK, MessageBoxIcon.Error)
			Exit Sub
		Catch ex As Exception
			MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End Try
		conn.Close()
	End Sub

End Class