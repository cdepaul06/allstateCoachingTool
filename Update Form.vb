Option Strict On
Imports System.Data.OleDb

Public Class frmUpdate
	Private Sub frmUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		'Populate the comboboxes for managers.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Managers FROM Managers"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				cmbCurrentManager.Items.Add(myReader("Managers"))
				cmbRemoveManager.Items.Add(myReader("Managers"))
				cmbTransferManager.Items.Add(myReader("Managers"))
				cmbToManager.Items.Add(myReader("Managers"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub cmbCurrentManager_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCurrentManager.SelectedIndexChanged
		'Populate the employee combobox based off the selected manager.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Employees FROM Employees WHERE Manager='" & cmbCurrentManager.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			cmbCurrentEmployee.Items.Clear()
			cmbRemoveEmployee.Items.Clear()

			While myReader.Read
				cmbCurrentEmployee.Items.Add(myReader("Employees"))
				cmbRemoveEmployee.Items.Add(myReader("Employees"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DeleteManager()
		Dim query As String = String.Empty
		query &= "DELETE FROM Managers WHERE Managers ='" & cmbRemoveManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
				End With
				Try
					conn.Open()
					comm.ExecuteNonQuery()
					conn.Close()
				Catch ex As OleDbException
					MessageBox.Show(ex.Message.ToString(), "Error.")
				End Try
			End Using
		End Using
	End Sub

	Private Sub AddManager()
		'Declare variables
		Dim strManager = CStr(txtAddManager.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO Managers (Managers) VALUES (@Managers)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Managers", strManager)
				End With
				Try
					conn.Open()
					comm.ExecuteNonQuery()
					conn.Close()
				Catch ex As OleDbException
					MessageBox.Show(ex.Message.ToString(), "Error.")
				End Try
			End Using
		End Using
	End Sub

	Private Sub AddEmployee()
		'Declare variables
		Dim strManager = CStr(cmbToManager.SelectedItem.ToString)
		Dim strEmployee = CStr(txtAddEmployee.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO Employees (Manager, Employees) VALUES (@Manager, Employees)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employees", strEmployee)
				End With
				Try
					conn.Open()
					comm.ExecuteNonQuery()
					conn.Close()
				Catch ex As OleDbException
					MessageBox.Show(ex.Message.ToString(), "Error.")
				End Try
			End Using
		End Using

		cmbCurrentEmployee.Items.Clear()
		cmbCurrentManager.Items.Clear()
		cmbRemoveManager.Items.Clear()
		cmbTransferManager.Items.Clear()
		cmbToManager.Items.Clear()
	End Sub

	Private Sub DeleteEmployee()
		Dim query As String = String.Empty
		query &= "DELETE FROM Employees WHERE Employees ='" & cmbRemoveEmployee.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
				End With
				Try
					conn.Open()
					comm.ExecuteNonQuery()
					conn.Close()
				Catch ex As OleDbException
					MessageBox.Show(ex.Message.ToString(), "Error.")
				End Try
			End Using
		End Using

		cmbCurrentEmployee.Items.Clear()
		cmbCurrentManager.Items.Clear()
		cmbRemoveManager.Items.Clear()
		cmbTransferManager.Items.Clear()
		cmbToManager.Items.Clear()

	End Sub

	Private Sub TransferEmployee()
		Dim strManager = cmbToManager.SelectedItem.ToString
		Dim strEmployee = cmbCurrentEmployee.SelectedItem.ToString
		Dim query As String = String.Empty
		query &= "UPDATE Employees SET Manager = @Manager, Employees = @Employees WHERE Employees ='" & cmbCurrentEmployee.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employees", strEmployee)
				End With
				Try
					conn.Open()
					comm.ExecuteNonQuery()
					conn.Close()
				Catch ex As OleDbException
					MessageBox.Show(ex.Message.ToString(), "Error.")
				End Try
			End Using
		End Using

		cmbCurrentEmployee.Items.Clear()
		cmbCurrentManager.Items.Clear()
		cmbRemoveManager.Items.Clear()
		cmbTransferManager.Items.Clear()
		cmbToManager.Items.Clear()

	End Sub

	Private Sub btnManagerUpdate_Click(sender As Object, e As EventArgs) Handles btnAddManager.Click
		'Add the name provided to the database of managers, reload the current managers and then display a messagebox.
		AddManager()
		'Populate the comboboxes for managers.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Managers FROM Managers"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			cmbCurrentManager.Items.Clear()
			cmbRemoveManager.Items.Clear()
			cmbTransferManager.Items.Clear()
			cmbToManager.Items.Clear()

			While myReader.Read
				cmbCurrentManager.Items.Add(myReader("Managers"))
				cmbRemoveManager.Items.Add(myReader("Managers"))
				cmbTransferManager.Items.Add(myReader("Managers"))
				cmbToManager.Items.Add(myReader("Managers"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
		MessageBox.Show("Manager has been added.", "Admin Manager", MessageBoxButtons.OK, MessageBoxIcon.Information)
		txtAddManager.Clear()

	End Sub

	Private Sub btnRemoveManager_Click(sender As Object, e As EventArgs) Handles btnRemoveManager.Click
		'Remove the selected name from the database of managers, reload the current managers and then display a messagebox.
		DeleteManager()
		'Populate the comboboxes for managers.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Managers FROM Managers"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			cmbCurrentManager.Items.Clear()
			cmbRemoveManager.Items.Clear()
			cmbTransferManager.Items.Clear()
			cmbToManager.Items.Clear()

			While myReader.Read
				cmbCurrentManager.Items.Add(myReader("Managers"))
				cmbRemoveManager.Items.Add(myReader("Managers"))
				cmbTransferManager.Items.Add(myReader("Managers"))
				cmbToManager.Items.Add(myReader("Managers"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
		MessageBox.Show("Manager has been removed.", "Admin Manager", MessageBoxButtons.OK, MessageBoxIcon.Information)
	End Sub

	Private Sub btnAddEmployee_Click(sender As Object, e As EventArgs) Handles btnAddEmployee.Click
		'Use the add employee subprocedure to add a new employee and assign them to a manager; display a messagebox informing user of change.
		AddEmployee()
		MessageBox.Show("Employee has been added.", "Admin Employee", MessageBoxButtons.OK, MessageBoxIcon.Information)
		cmbCurrentEmployee.Items.Clear()
		cmbCurrentManager.Items.Clear()
		txtAddEmployee.Clear()

		'Populate the comboboxes for managers.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Managers FROM Managers"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				cmbCurrentManager.Items.Add(myReader("Managers"))
				cmbRemoveManager.Items.Add(myReader("Managers"))
				cmbTransferManager.Items.Add(myReader("Managers"))
				cmbToManager.Items.Add(myReader("Managers"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try

	End Sub

	Private Sub btnRemoveEmp_Click(sender As Object, e As EventArgs) Handles btnRemoveEmp.Click
		'Use the remove employee subprocedure to remove an employee; display a messagebox informing user of change.
		DeleteEmployee()
		MessageBox.Show("Employee has been removed.", "Admin Employee", MessageBoxButtons.OK, MessageBoxIcon.Information)
		cmbCurrentEmployee.Items.Clear()
		cmbCurrentManager.Items.Clear()

		'Populate the comboboxes for managers.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Managers FROM Managers"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				cmbCurrentManager.Items.Add(myReader("Managers"))
				cmbRemoveManager.Items.Add(myReader("Managers"))
				cmbTransferManager.Items.Add(myReader("Managers"))
				cmbToManager.Items.Add(myReader("Managers"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try

	End Sub

	Private Sub btnTransferEmp_Click(sender As Object, e As EventArgs) Handles btnTransferEmp.Click
		'Use the transfer employee subprocedure to transfer an employee; display a messagebox informing user of change.
		TransferEmployee()
		MessageBox.Show("Employee has been transfered.", "Admin Employee", MessageBoxButtons.OK, MessageBoxIcon.Information)
		cmbCurrentEmployee.Items.Clear()
		cmbCurrentManager.Items.Clear()

		'Populate the comboboxes for managers.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Managers FROM Managers"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				cmbCurrentManager.Items.Add(myReader("Managers"))
				cmbRemoveManager.Items.Add(myReader("Managers"))
				cmbTransferManager.Items.Add(myReader("Managers"))
				cmbToManager.Items.Add(myReader("Managers"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try

	End Sub

	Private Sub cmbTransferManager_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTransferManager.SelectedIndexChanged
		'Populate the employee combobox based off the selected manager.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Employees FROM Employees WHERE Manager='" & cmbTransferManager.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			cmbCurrentEmployee.Items.Clear()
			cmbRemoveEmployee.Items.Clear()

			While myReader.Read
				cmbCurrentEmployee.Items.Add(myReader("Employees"))
				cmbRemoveEmployee.Items.Add(myReader("Employees"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub
End Class