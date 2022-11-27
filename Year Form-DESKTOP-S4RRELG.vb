Option Strict On
Imports System.Data.OleDb

Public Class frmYear
	Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbManager.SelectedIndexChanged
		'Populate the employee combobox based off the selected manager.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Employees FROM Employees WHERE Manager=" & "'" & cmbManager.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			cmbEmployee.Items.Clear()
			While myReader.Read
				cmbEmployee.Items.Add(myReader("Employees"))
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub frmYear_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		'Populate the comboboxes for managers and months.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Managers FROM Managers"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				cmbManager.Items.Add(myReader("Managers")).ToString()
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
		Me.Close()
	End Sub

	'Display the metrics for the currently selected employee.
	Private Sub DisplayJanMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM January WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblJanDuration.Text = myReader("Duration").ToString
				lblJanCSAT.Text = myReader("CSAT").ToString
				lblJanAway.Text = myReader("Away").ToString
				lblJanQual.Text = myReader("Quality").ToString
				lblJanDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayFebMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM February WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblFebDuration.Text = myReader("Duration").ToString
				lblFebCSAT.Text = myReader("CSAT").ToString
				lblFebAway.Text = myReader("Away").ToString
				lblFebQual.Text = myReader("Quality").ToString
				lblFebDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayMarMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM March WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblMarDuration.Text = myReader("Duration").ToString
				lblMarCSAT.Text = myReader("CSAT").ToString
				lblMarAway.Text = myReader("Away").ToString
				lblMarQual.Text = myReader("Quality").ToString
				lblMarDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayAprMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM April WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblAprDuration.Text = myReader("Duration").ToString
				lblAprCSAT.Text = myReader("CSAT").ToString
				lblAprAway.Text = myReader("Away").ToString
				lblAprQual.Text = myReader("Quality").ToString
				lblAprDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayMayMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM May WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblMayDuration.Text = myReader("Duration").ToString
				lblMayCSAT.Text = myReader("CSAT").ToString
				lblMayAway.Text = myReader("Away").ToString
				lblMayQual.Text = myReader("Quality").ToString
				lblMayDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayJunMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM June WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblJunDuration.Text = myReader("Duration").ToString
				lblJunCSAT.Text = myReader("CSAT").ToString
				lblJunAway.Text = myReader("Away").ToString
				lblJunQual.Text = myReader("Quality").ToString
				lblJunDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayJulMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM July WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblJulDuration.Text = myReader("Duration").ToString
				lblJulCSAT.Text = myReader("CSAT").ToString
				lblJulAway.Text = myReader("Away").ToString
				lblJulQual.Text = myReader("Quality").ToString
				lblJulDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayAugMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM August WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblAugDuration.Text = myReader("Duration").ToString
				lblAugCSAT.Text = myReader("CSAT").ToString
				lblAugAway.Text = myReader("Away").ToString
				lblAugQual.Text = myReader("Quality").ToString
				lblAugDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplaySepMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM September WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblSepDuration.Text = myReader("Duration").ToString
				lblSepCSAT.Text = myReader("CSAT").ToString
				lblSepAway.Text = myReader("Away").ToString
				lblSepQual.Text = myReader("Quality").ToString
				lblSepDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayOctMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM October WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblOctDuration.Text = myReader("Duration").ToString
				lblOctCSAT.Text = myReader("CSAT").ToString
				lblOctAway.Text = myReader("Away").ToString
				lblOctQual.Text = myReader("Quality").ToString
				lblOctDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayNovMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM November WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblNovDuration.Text = myReader("Duration").ToString
				lblNovCSAT.Text = myReader("CSAT").ToString
				lblNovAway.Text = myReader("Away").ToString
				lblNovQual.Text = myReader("Quality").ToString
				lblNovDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DisplayDecMetrics()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Duration, CSAT, Away, Quality, Development FROM December WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblDecDuration.Text = myReader("Duration").ToString
				lblDecCSAT.Text = myReader("CSAT").ToString
				lblDecAway.Text = myReader("Away").ToString
				lblDecQual.Text = myReader("Quality").ToString
				lblDecDev.Text = myReader("Development").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub ColorJanMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblJanDuration.Text <= "16:00" Then
			lblJanDuration.BackColor = Color.SteelBlue
			lblJanDuration.ForeColor = Color.White
		End If

		If lblJanDuration.Text > "16:00" Then
			lblJanDuration.BackColor = Color.ForestGreen
			lblJanDuration.ForeColor = Color.White
		End If

		If lblJanDuration.Text > "21:00" Then
			lblJanDuration.BackColor = Color.Gold
			lblJanDuration.ForeColor = Color.Black
		End If

		If lblJanDuration.Text > "25:00" Then
			lblJanDuration.BackColor = Color.IndianRed
			lblJanDuration.ForeColor = Color.Black
		End If

		If lblJanCSAT.Text >= "87.00" Then
			lblJanCSAT.BackColor = Color.SteelBlue
			lblJanCSAT.ForeColor = Color.White
		End If

		If lblJanCSAT.Text <= "86.99" Then
			lblJanCSAT.BackColor = Color.ForestGreen
			lblJanCSAT.ForeColor = Color.White
		End If

		If lblJanCSAT.Text <= "74.99" Then
			lblJanCSAT.BackColor = Color.Gold
			lblJanCSAT.ForeColor = Color.Black
		End If

		If lblJanCSAT.Text < "70.00" Then
			lblJanCSAT.BackColor = Color.IndianRed
			lblJanCSAT.ForeColor = Color.Black
		End If

		If lblJanAway.Text <= "26.00" Then
			lblJanAway.BackColor = Color.SteelBlue
			lblJanAway.ForeColor = Color.White
		End If

		If lblJanAway.Text > "26.00" Then
			lblJanAway.BackColor = Color.ForestGreen
			lblJanAway.ForeColor = Color.White
		End If

		If lblJanAway.Text > "29.99" Then
			lblJanAway.BackColor = Color.Gold
			lblJanAway.ForeColor = Color.Black
		End If

		If lblJanAway.Text > "37.00" Then
			lblJanAway.BackColor = Color.IndianRed
			lblJanAway.ForeColor = Color.Black
		End If

		If lblJanDuration.Text Is "" Then
			lblJanDuration.BackColor = Color.White
		End If

		If lblJanCSAT.Text Is "" Then
			lblJanCSAT.BackColor = Color.White
		End If

		If lblJanAway.Text Is "" Then
			lblJanAway.BackColor = Color.White
		End If

		If lblJanQual.Text Is "" Then
			lblJanQual.BackColor = Color.White
		End If

		If lblJanDev.Text Is "" Then
			lblJanDev.BackColor = Color.White
		End If

		If lblJanCSAT.Text = "100" Then
			lblJanCSAT.BackColor = Color.SteelBlue
			lblJanCSAT.ForeColor = Color.White
		End If

		If lblJanAway.Text = "100" Then
			lblJanAway.BackColor = Color.IndianRed
			lblJanAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorFebMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblFebDuration.Text <= "16:00" Then
			lblFebDuration.BackColor = Color.SteelBlue
			lblFebDuration.ForeColor = Color.White
		End If

		If lblFebDuration.Text > "16:00" Then
			lblFebDuration.BackColor = Color.ForestGreen
			lblFebDuration.ForeColor = Color.White
		End If

		If lblFebDuration.Text > "21:00" Then
			lblFebDuration.BackColor = Color.Gold
			lblFebDuration.ForeColor = Color.Black
		End If

		If lblFebDuration.Text > "25:00" Then
			lblFebDuration.BackColor = Color.IndianRed
			lblFebDuration.ForeColor = Color.Black
		End If

		If lblFebCSAT.Text >= "87.00" Then
			lblFebCSAT.BackColor = Color.SteelBlue
			lblFebCSAT.ForeColor = Color.White
		End If

		If lblFebCSAT.Text <= "86.99" Then
			lblFebCSAT.BackColor = Color.ForestGreen
			lblFebCSAT.ForeColor = Color.White
		End If

		If lblFebCSAT.Text <= "74.99" Then
			lblFebCSAT.BackColor = Color.Gold
			lblFebCSAT.ForeColor = Color.Black
		End If

		If lblFebCSAT.Text < "70.00" Then
			lblFebCSAT.BackColor = Color.IndianRed
			lblFebCSAT.ForeColor = Color.Black
		End If

		If lblFebAway.Text <= "26.00" Then
			lblFebAway.BackColor = Color.SteelBlue
			lblFebAway.ForeColor = Color.White
		End If

		If lblFebAway.Text > "26.00" Then
			lblFebAway.BackColor = Color.ForestGreen
			lblFebAway.ForeColor = Color.White
		End If

		If lblFebAway.Text > "29.99" Then
			lblFebAway.BackColor = Color.Gold
			lblFebAway.ForeColor = Color.Black
		End If

		If lblFebAway.Text > "37.00" Then
			lblFebAway.BackColor = Color.IndianRed
			lblFebAway.ForeColor = Color.Black
		End If

		If lblFebDuration.Text Is "" Then
			lblFebDuration.BackColor = Color.White
		End If

		If lblFebCSAT.Text Is "" Then
			lblFebCSAT.BackColor = Color.White
		End If

		If lblFebAway.Text Is "" Then
			lblFebAway.BackColor = Color.White
		End If

		If lblFebQual.Text Is "" Then
			lblFebQual.BackColor = Color.White
		End If

		If lblFebDev.Text Is "" Then
			lblFebDev.BackColor = Color.White
		End If

		If lblFebCSAT.Text = "100" Then
			lblFebCSAT.BackColor = Color.SteelBlue
			lblFebCSAT.ForeColor = Color.White
		End If

		If lblFebAway.Text = "100" Then
			lblFebAway.BackColor = Color.IndianRed
			lblFebAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorMarMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblMarDuration.Text <= "16:00" Then
			lblMarDuration.BackColor = Color.SteelBlue
			lblMarDuration.ForeColor = Color.White
		End If

		If lblMarDuration.Text > "16:00" Then
			lblMarDuration.BackColor = Color.ForestGreen
			lblMarDuration.ForeColor = Color.White
		End If

		If lblMarDuration.Text > "21:00" Then
			lblMarDuration.BackColor = Color.Gold
			lblMarDuration.ForeColor = Color.Black
		End If

		If lblMarDuration.Text > "25:00" Then
			lblMarDuration.BackColor = Color.IndianRed
			lblMarDuration.ForeColor = Color.Black
		End If

		If lblMarCSAT.Text >= "87.00" Then
			lblMarCSAT.BackColor = Color.SteelBlue
			lblMarCSAT.ForeColor = Color.White
		End If

		If lblMarCSAT.Text <= "86.99" Then
			lblMarCSAT.BackColor = Color.ForestGreen
			lblMarCSAT.ForeColor = Color.White
		End If

		If lblMarCSAT.Text <= "74.99" Then
			lblMarCSAT.BackColor = Color.Gold
			lblMarCSAT.ForeColor = Color.Black
		End If

		If lblMarCSAT.Text < "70.00" Then
			lblMarCSAT.BackColor = Color.IndianRed
			lblMarCSAT.ForeColor = Color.Black
		End If

		If lblMarAway.Text <= "26.00" Then
			lblMarAway.BackColor = Color.SteelBlue
			lblMarAway.ForeColor = Color.White
		End If

		If lblMarAway.Text > "26.00" Then
			lblMarAway.BackColor = Color.ForestGreen
			lblMarAway.ForeColor = Color.White
		End If

		If lblMarAway.Text > "29.99" Then
			lblMarAway.BackColor = Color.Gold
			lblMarAway.ForeColor = Color.Black
		End If

		If lblMarAway.Text > "37.00" Then
			lblMarAway.BackColor = Color.IndianRed
			lblMarAway.ForeColor = Color.Black
		End If

		If lblMarDuration.Text Is "" Then
			lblMarDuration.BackColor = Color.White
		End If

		If lblMarCSAT.Text Is "" Then
			lblMarCSAT.BackColor = Color.White
		End If

		If lblMarAway.Text Is "" Then
			lblMarAway.BackColor = Color.White
		End If

		If lblMarQual.Text Is "" Then
			lblMarQual.BackColor = Color.White
		End If

		If lblMarDev.Text Is "" Then
			lblMarDev.BackColor = Color.White
		End If

		If lblMarCSAT.Text = "100" Then
			lblMarCSAT.BackColor = Color.SteelBlue
			lblMarCSAT.ForeColor = Color.White
		End If

		If lblMarAway.Text = "100" Then
			lblMarAway.BackColor = Color.IndianRed
			lblMarAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorAprMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblAprDuration.Text <= "16:00" Then
			lblAprDuration.BackColor = Color.SteelBlue
			lblAprDuration.ForeColor = Color.White
		End If

		If lblAprDuration.Text > "16:00" Then
			lblAprDuration.BackColor = Color.ForestGreen
			lblAprDuration.ForeColor = Color.White
		End If

		If lblAprDuration.Text > "21:00" Then
			lblAprDuration.BackColor = Color.Gold
			lblAprDuration.ForeColor = Color.Black
		End If

		If lblAprDuration.Text > "25:00" Then
			lblAprDuration.BackColor = Color.IndianRed
			lblAprDuration.ForeColor = Color.Black
		End If

		If lblAprCSAT.Text >= "87.00" Then
			lblAprCSAT.BackColor = Color.SteelBlue
			lblAprCSAT.ForeColor = Color.White
		End If

		If lblAprCSAT.Text <= "86.99" Then
			lblAprCSAT.BackColor = Color.ForestGreen
			lblAprCSAT.ForeColor = Color.White
		End If

		If lblAprCSAT.Text <= "74.99" Then
			lblAprCSAT.BackColor = Color.Gold
			lblAprCSAT.ForeColor = Color.Black
		End If

		If lblAprCSAT.Text < "70.00" Then
			lblAprCSAT.BackColor = Color.IndianRed
			lblAprCSAT.ForeColor = Color.Black
		End If

		If lblAprAway.Text <= "26.00" Then
			lblAprAway.BackColor = Color.SteelBlue
			lblAprAway.ForeColor = Color.White
		End If

		If lblAprAway.Text > "26.00" Then
			lblAprAway.BackColor = Color.ForestGreen
			lblAprAway.ForeColor = Color.White
		End If

		If lblAprAway.Text > "29.99" Then
			lblAprAway.BackColor = Color.Gold
			lblAprAway.ForeColor = Color.Black
		End If

		If lblAprAway.Text > "37.00" Then
			lblAprAway.BackColor = Color.IndianRed
			lblAprAway.ForeColor = Color.Black
		End If

		If lblAprDuration.Text Is "" Then
			lblAprDuration.BackColor = Color.White
		End If

		If lblAprCSAT.Text Is "" Then
			lblAprCSAT.BackColor = Color.White
		End If

		If lblAprAway.Text Is "" Then
			lblAprAway.BackColor = Color.White
		End If

		If lblAprQual.Text Is "" Then
			lblAprQual.BackColor = Color.White
		End If

		If lblAprDev.Text Is "" Then
			lblAprDev.BackColor = Color.White
		End If

		If lblAprCSAT.Text = "100" Then
			lblAprCSAT.BackColor = Color.SteelBlue
			lblAprCSAT.ForeColor = Color.White
		End If

		If lblAprAway.Text = "100" Then
			lblAprAway.BackColor = Color.IndianRed
			lblAprAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorMayMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblMayDuration.Text <= "16:00" Then
			lblMayDuration.BackColor = Color.SteelBlue
			lblMayDuration.ForeColor = Color.White
		End If

		If lblMayDuration.Text > "16:00" Then
			lblMayDuration.BackColor = Color.ForestGreen
			lblMayDuration.ForeColor = Color.White
		End If

		If lblMayDuration.Text > "21:00" Then
			lblMayDuration.BackColor = Color.Gold
			lblMayDuration.ForeColor = Color.Black
		End If

		If lblMayDuration.Text > "25:00" Then
			lblMayDuration.BackColor = Color.IndianRed
			lblMayDuration.ForeColor = Color.Black
		End If

		If lblMayCSAT.Text >= "87.00" Then
			lblMayCSAT.BackColor = Color.SteelBlue
			lblMayCSAT.ForeColor = Color.White
		End If

		If lblMayCSAT.Text <= "86.99" Then
			lblMayCSAT.BackColor = Color.ForestGreen
			lblMayCSAT.ForeColor = Color.White
		End If

		If lblMayCSAT.Text <= "74.99" Then
			lblMayCSAT.BackColor = Color.Gold
			lblMayCSAT.ForeColor = Color.Black
		End If

		If lblMayCSAT.Text < "70.00" Then
			lblMayCSAT.BackColor = Color.IndianRed
			lblMayCSAT.ForeColor = Color.Black
		End If

		If lblMayAway.Text <= "26.00" Then
			lblMayAway.BackColor = Color.SteelBlue
			lblMayAway.ForeColor = Color.White
		End If

		If lblMayAway.Text > "26.00" Then
			lblMayAway.BackColor = Color.ForestGreen
			lblMayAway.ForeColor = Color.White
		End If

		If lblMayAway.Text > "29.99" Then
			lblMayAway.BackColor = Color.Gold
			lblMayAway.ForeColor = Color.Black
		End If

		If lblMayAway.Text > "37.00" Then
			lblMayAway.BackColor = Color.IndianRed
			lblMayAway.ForeColor = Color.Black
		End If

		If lblMayDuration.Text Is "" Then
			lblMayDuration.BackColor = Color.White
		End If

		If lblMayCSAT.Text Is "" Then
			lblMayCSAT.BackColor = Color.White
		End If

		If lblMayAway.Text Is "" Then
			lblMayAway.BackColor = Color.White
		End If

		If lblMayQual.Text Is "" Then
			lblMayQual.BackColor = Color.White
		End If

		If lblMayDev.Text Is "" Then
			lblMayDev.BackColor = Color.White
		End If

		If lblMayCSAT.Text = "100" Then
			lblMayCSAT.BackColor = Color.SteelBlue
			lblMayCSAT.ForeColor = Color.White
		End If

		If lblMayAway.Text = "100" Then
			lblMayAway.BackColor = Color.IndianRed
			lblMayAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorJunMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblJunDuration.Text <= "16:00" Then
			lblJunDuration.BackColor = Color.SteelBlue
			lblJunDuration.ForeColor = Color.White
		End If

		If lblJunDuration.Text > "16:00" Then
			lblJunDuration.BackColor = Color.ForestGreen
			lblJunDuration.ForeColor = Color.White
		End If

		If lblJunDuration.Text > "21:00" Then
			lblJunDuration.BackColor = Color.Gold
			lblJunDuration.ForeColor = Color.Black
		End If

		If lblJunDuration.Text > "25:00" Then
			lblJunDuration.BackColor = Color.IndianRed
			lblJunDuration.ForeColor = Color.Black
		End If

		If lblJunCSAT.Text >= "87.00" Then
			lblJunCSAT.BackColor = Color.SteelBlue
			lblJunCSAT.ForeColor = Color.White
		End If

		If lblJunCSAT.Text <= "86.99" Then
			lblJunCSAT.BackColor = Color.ForestGreen
			lblJunCSAT.ForeColor = Color.White
		End If

		If lblJunCSAT.Text <= "74.99" Then
			lblJunCSAT.BackColor = Color.Gold
			lblJunCSAT.ForeColor = Color.Black
		End If

		If lblJunCSAT.Text < "70.00" Then
			lblJunCSAT.BackColor = Color.IndianRed
			lblJunCSAT.ForeColor = Color.Black
		End If

		If lblJunAway.Text <= "26.00" Then
			lblJunAway.BackColor = Color.SteelBlue
			lblJunAway.ForeColor = Color.White
		End If

		If lblJunAway.Text > "26.00" Then
			lblJunAway.BackColor = Color.ForestGreen
			lblJunAway.ForeColor = Color.White
		End If

		If lblJunAway.Text > "29.99" Then
			lblJunAway.BackColor = Color.Gold
			lblJunAway.ForeColor = Color.Black
		End If

		If lblJunAway.Text > "37.00" Then
			lblJunAway.BackColor = Color.IndianRed
			lblJunAway.ForeColor = Color.Black
		End If

		If lblJunDuration.Text Is "" Then
			lblJunDuration.BackColor = Color.White
		End If

		If lblJunCSAT.Text Is "" Then
			lblJunCSAT.BackColor = Color.White
		End If

		If lblJunAway.Text Is "" Then
			lblJunAway.BackColor = Color.White
		End If

		If lblJunQual.Text Is "" Then
			lblJunQual.BackColor = Color.White
		End If

		If lblJunDev.Text Is "" Then
			lblJunDev.BackColor = Color.White
		End If

		If lblJunCSAT.Text = "100" Then
			lblJunCSAT.BackColor = Color.SteelBlue
			lblJunCSAT.ForeColor = Color.White
		End If

		If lblJunAway.Text = "100" Then
			lblJunAway.BackColor = Color.IndianRed
			lblJunAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorJulMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblJulDuration.Text <= "16:00" Then
			lblJulDuration.BackColor = Color.SteelBlue
			lblJulDuration.ForeColor = Color.White
		End If

		If lblJulDuration.Text > "16:00" Then
			lblJulDuration.BackColor = Color.ForestGreen
			lblJulDuration.ForeColor = Color.White
		End If

		If lblJulDuration.Text > "21:00" Then
			lblJulDuration.BackColor = Color.Gold
			lblJulDuration.ForeColor = Color.Black
		End If

		If lblJulDuration.Text > "25:00" Then
			lblJulDuration.BackColor = Color.IndianRed
			lblJulDuration.ForeColor = Color.Black
		End If

		If lblJulCSAT.Text >= "87.00" Then
			lblJulCSAT.BackColor = Color.SteelBlue
			lblJulCSAT.ForeColor = Color.White
		End If

		If lblJulCSAT.Text <= "86.99" Then
			lblJulCSAT.BackColor = Color.ForestGreen
			lblJulCSAT.ForeColor = Color.White
		End If

		If lblJulCSAT.Text <= "74.99" Then
			lblJulCSAT.BackColor = Color.Gold
			lblJulCSAT.ForeColor = Color.Black
		End If

		If lblJulCSAT.Text < "70.00" Then
			lblJulCSAT.BackColor = Color.IndianRed
			lblJulCSAT.ForeColor = Color.Black
		End If

		If lblJulAway.Text <= "26.00" Then
			lblJulAway.BackColor = Color.SteelBlue
			lblJulAway.ForeColor = Color.White
		End If

		If lblJulAway.Text > "26.00" Then
			lblJulAway.BackColor = Color.ForestGreen
			lblJulAway.ForeColor = Color.White
		End If

		If lblJulAway.Text > "29.99" Then
			lblJulAway.BackColor = Color.Gold
			lblJulAway.ForeColor = Color.Black
		End If

		If lblJulAway.Text > "37.00" Then
			lblJulAway.BackColor = Color.IndianRed
			lblJulAway.ForeColor = Color.Black
		End If

		If lblJulDuration.Text Is "" Then
			lblJulDuration.BackColor = Color.White
		End If

		If lblJulCSAT.Text Is "" Then
			lblJulCSAT.BackColor = Color.White
		End If

		If lblJulAway.Text Is "" Then
			lblJulAway.BackColor = Color.White
		End If

		If lblJulQual.Text Is "" Then
			lblJulQual.BackColor = Color.White
		End If

		If lblJulDev.Text Is "" Then
			lblJulDev.BackColor = Color.White
		End If

		If lblJulCSAT.Text = "100" Then
			lblJulCSAT.BackColor = Color.SteelBlue
			lblJulCSAT.ForeColor = Color.White
		End If

		If lblJulAway.Text = "100" Then
			lblJulAway.BackColor = Color.IndianRed
			lblJulAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorAugMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblAugDuration.Text <= "16:00" Then
			lblAugDuration.BackColor = Color.SteelBlue
			lblAugDuration.ForeColor = Color.White
		End If

		If lblAugDuration.Text > "16:00" Then
			lblAugDuration.BackColor = Color.ForestGreen
			lblAugDuration.ForeColor = Color.White
		End If

		If lblAugDuration.Text > "21:00" Then
			lblAugDuration.BackColor = Color.Gold
			lblAugDuration.ForeColor = Color.Black
		End If

		If lblAugDuration.Text > "25:00" Then
			lblAugDuration.BackColor = Color.IndianRed
			lblAugDuration.ForeColor = Color.Black
		End If

		If lblAugCSAT.Text >= "87.00" Then
			lblAugCSAT.BackColor = Color.SteelBlue
			lblAugCSAT.ForeColor = Color.White
		End If

		If lblAugCSAT.Text <= "86.99" Then
			lblAugCSAT.BackColor = Color.ForestGreen
			lblAugCSAT.ForeColor = Color.White
		End If

		If lblAugCSAT.Text <= "74.99" Then
			lblAugCSAT.BackColor = Color.Gold
			lblAugCSAT.ForeColor = Color.Black
		End If

		If lblAugCSAT.Text < "70.00" Then
			lblAugCSAT.BackColor = Color.IndianRed
			lblAugCSAT.ForeColor = Color.Black
		End If

		If lblAugAway.Text <= "26.00" Then
			lblAugAway.BackColor = Color.SteelBlue
			lblAugAway.ForeColor = Color.White
		End If

		If lblAugAway.Text > "26.00" Then
			lblAugAway.BackColor = Color.ForestGreen
			lblAugAway.ForeColor = Color.White
		End If

		If lblAugAway.Text > "29.99" Then
			lblAugAway.BackColor = Color.Gold
			lblAugAway.ForeColor = Color.Black
		End If

		If lblAugAway.Text > "37.00" Then
			lblAugAway.BackColor = Color.IndianRed
			lblAugAway.ForeColor = Color.Black
		End If

		If lblAugDuration.Text Is "" Then
			lblAugDuration.BackColor = Color.White
		End If

		If lblAugCSAT.Text Is "" Then
			lblAugCSAT.BackColor = Color.White
		End If

		If lblAugAway.Text Is "" Then
			lblAugAway.BackColor = Color.White
		End If

		If lblAugQual.Text Is "" Then
			lblAugQual.BackColor = Color.White
		End If

		If lblAugDev.Text Is "" Then
			lblAugDev.BackColor = Color.White
		End If

		If lblAugCSAT.Text = "100" Then
			lblAugCSAT.BackColor = Color.SteelBlue
			lblAugCSAT.ForeColor = Color.White
		End If

		If lblAugAway.Text = "100" Then
			lblAugAway.BackColor = Color.IndianRed
			lblAugAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorSepMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblSepDuration.Text <= "16:00" Then
			lblSepDuration.BackColor = Color.SteelBlue
			lblSepDuration.ForeColor = Color.White
		End If

		If lblSepDuration.Text > "16:00" Then
			lblSepDuration.BackColor = Color.ForestGreen
			lblSepDuration.ForeColor = Color.White
		End If

		If lblSepDuration.Text > "21:00" Then
			lblSepDuration.BackColor = Color.Gold
			lblSepDuration.ForeColor = Color.Black
		End If

		If lblSepDuration.Text > "25:00" Then
			lblSepDuration.BackColor = Color.IndianRed
			lblSepDuration.ForeColor = Color.Black
		End If

		If lblSepCSAT.Text >= "87.00" Then
			lblSepCSAT.BackColor = Color.SteelBlue
			lblSepCSAT.ForeColor = Color.White
		End If

		If lblSepCSAT.Text <= "86.99" Then
			lblSepCSAT.BackColor = Color.ForestGreen
			lblSepCSAT.ForeColor = Color.White
		End If

		If lblSepCSAT.Text <= "74.99" Then
			lblSepCSAT.BackColor = Color.Gold
			lblSepCSAT.ForeColor = Color.Black
		End If

		If lblSepCSAT.Text < "70.00" Then
			lblSepCSAT.BackColor = Color.IndianRed
			lblSepCSAT.ForeColor = Color.Black
		End If

		If lblSepAway.Text <= "26.00" Then
			lblSepAway.BackColor = Color.SteelBlue
			lblSepAway.ForeColor = Color.White
		End If

		If lblSepAway.Text > "26.00" Then
			lblSepAway.BackColor = Color.ForestGreen
			lblSepAway.ForeColor = Color.White
		End If

		If lblSepAway.Text > "29.99" Then
			lblSepAway.BackColor = Color.Gold
			lblSepAway.ForeColor = Color.Black
		End If

		If lblSepAway.Text > "37.00" Then
			lblSepAway.BackColor = Color.IndianRed
			lblSepAway.ForeColor = Color.Black
		End If

		If lblSepDuration.Text Is "" Then
			lblSepDuration.BackColor = Color.White
		End If

		If lblSepCSAT.Text Is "" Then
			lblSepCSAT.BackColor = Color.White
		End If

		If lblSepAway.Text Is "" Then
			lblSepAway.BackColor = Color.White
		End If

		If lblSepQual.Text Is "" Then
			lblSepQual.BackColor = Color.White
		End If

		If lblSepDev.Text Is "" Then
			lblSepDev.BackColor = Color.White
		End If

		If lblSepCSAT.Text = "100" Then
			lblSepCSAT.BackColor = Color.SteelBlue
			lblSepCSAT.ForeColor = Color.White
		End If

		If lblSepAway.Text = "100" Then
			lblSepAway.BackColor = Color.IndianRed
			lblSepAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorOctMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblOctDuration.Text <= "16:00" Then
			lblOctDuration.BackColor = Color.SteelBlue
			lblOctDuration.ForeColor = Color.White
		End If

		If lblOctDuration.Text > "16:00" Then
			lblOctDuration.BackColor = Color.ForestGreen
			lblOctDuration.ForeColor = Color.White
		End If

		If lblOctDuration.Text > "21:00" Then
			lblOctDuration.BackColor = Color.Gold
			lblOctDuration.ForeColor = Color.Black
		End If

		If lblOctDuration.Text > "25:00" Then
			lblOctDuration.BackColor = Color.IndianRed
			lblOctDuration.ForeColor = Color.Black
		End If

		If lblOctCSAT.Text >= "87.00" Then
			lblOctCSAT.BackColor = Color.SteelBlue
			lblOctCSAT.ForeColor = Color.White
		End If

		If lblOctCSAT.Text <= "86.99" Then
			lblOctCSAT.BackColor = Color.ForestGreen
			lblOctCSAT.ForeColor = Color.White
		End If

		If lblOctCSAT.Text <= "74.99" Then
			lblOctCSAT.BackColor = Color.Gold
			lblOctCSAT.ForeColor = Color.Black
		End If

		If lblOctCSAT.Text < "70.00" Then
			lblOctCSAT.BackColor = Color.IndianRed
			lblOctCSAT.ForeColor = Color.Black
		End If

		If lblOctAway.Text <= "26.00" Then
			lblOctAway.BackColor = Color.SteelBlue
			lblOctAway.ForeColor = Color.White
		End If

		If lblOctAway.Text > "26.00" Then
			lblOctAway.BackColor = Color.ForestGreen
			lblOctAway.ForeColor = Color.White
		End If

		If lblOctAway.Text > "29.99" Then
			lblOctAway.BackColor = Color.Gold
			lblOctAway.ForeColor = Color.Black
		End If

		If lblOctAway.Text > "37.00" Then
			lblOctAway.BackColor = Color.IndianRed
			lblOctAway.ForeColor = Color.Black
		End If

		If lblOctDuration.Text Is "" Then
			lblOctDuration.BackColor = Color.White
		End If

		If lblOctCSAT.Text Is "" Then
			lblOctCSAT.BackColor = Color.White
		End If

		If lblOctAway.Text Is "" Then
			lblOctAway.BackColor = Color.White
		End If

		If lblOctQual.Text Is "" Then
			lblOctQual.BackColor = Color.White
		End If

		If lblOctDev.Text Is "" Then
			lblOctDev.BackColor = Color.White
		End If

		If lblOctCSAT.Text = "100" Then
			lblOctCSAT.BackColor = Color.SteelBlue
			lblOctCSAT.ForeColor = Color.White
		End If

		If lblOctAway.Text = "100" Then
			lblOctAway.BackColor = Color.IndianRed
			lblOctAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorNovMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblNovDuration.Text <= "16:00" Then
			lblNovDuration.BackColor = Color.SteelBlue
			lblNovDuration.ForeColor = Color.White
		End If

		If lblNovDuration.Text > "16:00" Then
			lblNovDuration.BackColor = Color.ForestGreen
			lblNovDuration.ForeColor = Color.White
		End If

		If lblNovDuration.Text > "21:00" Then
			lblNovDuration.BackColor = Color.Gold
			lblNovDuration.ForeColor = Color.Black
		End If

		If lblNovDuration.Text > "25:00" Then
			lblNovDuration.BackColor = Color.IndianRed
			lblNovDuration.ForeColor = Color.Black
		End If

		If lblNovCSAT.Text >= "87.00" Then
			lblNovCSAT.BackColor = Color.SteelBlue
			lblNovCSAT.ForeColor = Color.White
		End If

		If lblNovCSAT.Text <= "86.99" Then
			lblNovCSAT.BackColor = Color.ForestGreen
			lblNovCSAT.ForeColor = Color.White
		End If

		If lblNovCSAT.Text <= "74.99" Then
			lblNovCSAT.BackColor = Color.Gold
			lblNovCSAT.ForeColor = Color.Black
		End If

		If lblNovCSAT.Text < "70.00" Then
			lblNovCSAT.BackColor = Color.IndianRed
			lblNovCSAT.ForeColor = Color.Black
		End If

		If lblNovAway.Text <= "26.00" Then
			lblNovAway.BackColor = Color.SteelBlue
			lblNovAway.ForeColor = Color.White
		End If

		If lblNovAway.Text > "26.00" Then
			lblNovAway.BackColor = Color.ForestGreen
			lblNovAway.ForeColor = Color.White
		End If

		If lblNovAway.Text > "29.99" Then
			lblNovAway.BackColor = Color.Gold
			lblNovAway.ForeColor = Color.Black
		End If

		If lblNovAway.Text > "37.00" Then
			lblNovAway.BackColor = Color.IndianRed
			lblNovAway.ForeColor = Color.Black
		End If

		If lblNovDuration.Text Is "" Then
			lblNovDuration.BackColor = Color.White
		End If

		If lblNovCSAT.Text Is "" Then
			lblNovCSAT.BackColor = Color.White
		End If

		If lblNovAway.Text Is "" Then
			lblNovAway.BackColor = Color.White
		End If

		If lblNovQual.Text Is "" Then
			lblNovQual.BackColor = Color.White
		End If

		If lblNovDev.Text Is "" Then
			lblNovDev.BackColor = Color.White
		End If

		If lblNovCSAT.Text = "100" Then
			lblNovCSAT.BackColor = Color.SteelBlue
			lblNovCSAT.ForeColor = Color.White
		End If

		If lblNovAway.Text = "100" Then
			lblNovAway.BackColor = Color.IndianRed
			lblNovAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorDecMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblDecDuration.Text <= "16:00" Then
			lblDecDuration.BackColor = Color.SteelBlue
			lblDecDuration.ForeColor = Color.White
		End If

		If lblDecDuration.Text > "16:00" Then
			lblDecDuration.BackColor = Color.ForestGreen
			lblDecDuration.ForeColor = Color.White
		End If

		If lblDecDuration.Text > "21:00" Then
			lblDecDuration.BackColor = Color.Gold
			lblDecDuration.ForeColor = Color.Black
		End If

		If lblDecDuration.Text > "25:00" Then
			lblDecDuration.BackColor = Color.IndianRed
			lblDecDuration.ForeColor = Color.Black
		End If

		If lblDecCSAT.Text >= "87.00" Then
			lblDecCSAT.BackColor = Color.SteelBlue
			lblDecCSAT.ForeColor = Color.White
		End If

		If lblDecCSAT.Text <= "86.99" Then
			lblDecCSAT.BackColor = Color.ForestGreen
			lblDecCSAT.ForeColor = Color.White
		End If

		If lblDecCSAT.Text <= "74.99" Then
			lblDecCSAT.BackColor = Color.Gold
			lblDecCSAT.ForeColor = Color.Black
		End If

		If lblDecCSAT.Text < "70.00" Then
			lblDecCSAT.BackColor = Color.IndianRed
			lblDecCSAT.ForeColor = Color.Black
		End If

		If lblDecAway.Text <= "26.00" Then
			lblDecAway.BackColor = Color.SteelBlue
			lblDecAway.ForeColor = Color.White
		End If

		If lblDecAway.Text > "26.00" Then
			lblDecAway.BackColor = Color.ForestGreen
			lblDecAway.ForeColor = Color.White
		End If

		If lblDecAway.Text > "29.99" Then
			lblDecAway.BackColor = Color.Gold
			lblDecAway.ForeColor = Color.Black
		End If

		If lblDecAway.Text > "37.00" Then
			lblDecAway.BackColor = Color.IndianRed
			lblDecAway.ForeColor = Color.Black
		End If

		If lblDecDuration.Text Is "" Then
			lblDecDuration.BackColor = Color.White
		End If

		If lblDecCSAT.Text Is "" Then
			lblDecCSAT.BackColor = Color.White
		End If

		If lblDecAway.Text Is "" Then
			lblDecAway.BackColor = Color.White
		End If

		If lblDecQual.Text Is "" Then
			lblDecQual.BackColor = Color.White
		End If

		If lblDecDev.Text Is "" Then
			lblDecDev.BackColor = Color.White
		End If

		If lblDecCSAT.Text = "100" Then
			lblDecCSAT.BackColor = Color.SteelBlue
			lblDecCSAT.ForeColor = Color.White
		End If

		If lblDecAway.Text = "100" Then
			lblDecAway.BackColor = Color.IndianRed
			lblDecAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub ColorYearMetrics()
		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblYearDuration.Text <= "16:00" Then
			lblYearDuration.BackColor = Color.SteelBlue
			lblYearDuration.ForeColor = Color.White
		End If

		If lblYearDuration.Text > "16:00" Then
			lblYearDuration.BackColor = Color.ForestGreen
			lblYearDuration.ForeColor = Color.White
		End If

		If lblYearDuration.Text > "21:00" Then
			lblYearDuration.BackColor = Color.Gold
			lblYearDuration.ForeColor = Color.Black
		End If

		If lblYearDuration.Text > "25:00" Then
			lblYearDuration.BackColor = Color.IndianRed
			lblYearDuration.ForeColor = Color.Black
		End If

		If lblYearCSAT.Text >= "87.00" Then
			lblYearCSAT.BackColor = Color.SteelBlue
			lblYearCSAT.ForeColor = Color.White
		End If

		If lblYearCSAT.Text <= "86.99" Then
			lblYearCSAT.BackColor = Color.ForestGreen
			lblYearCSAT.ForeColor = Color.White
		End If

		If lblYearCSAT.Text <= "74.99" Then
			lblYearCSAT.BackColor = Color.Gold
			lblYearCSAT.ForeColor = Color.Black
		End If

		If lblYearCSAT.Text < "70.00" Then
			lblYearCSAT.BackColor = Color.IndianRed
			lblYearCSAT.ForeColor = Color.Black
		End If

		If lblYearAway.Text <= "26.00" Then
			lblYearAway.BackColor = Color.SteelBlue
			lblYearAway.ForeColor = Color.White
		End If

		If lblYearAway.Text > "26.00" Then
			lblYearAway.BackColor = Color.ForestGreen
			lblYearAway.ForeColor = Color.White
		End If

		If lblYearAway.Text > "29.99" Then
			lblYearAway.BackColor = Color.Gold
			lblYearAway.ForeColor = Color.Black
		End If

		If lblYearAway.Text > "37.00" Then
			lblYearAway.BackColor = Color.IndianRed
			lblYearAway.ForeColor = Color.Black
		End If

		If lblYearDuration.Text Is "" Then
			lblYearDuration.BackColor = Color.White
		End If

		If lblYearCSAT.Text Is "" Then
			lblYearCSAT.BackColor = Color.White
		End If

		If lblYearAway.Text Is "" Then
			lblYearAway.BackColor = Color.White
		End If

		If lblYearQual.Text Is "" Then
			lblYearQual.BackColor = Color.White
		End If

		If lblYearDev.Text Is "" Then
			lblYearDev.BackColor = Color.White
		End If

		If lblYearCSAT.Text = "100" Then
			lblYearCSAT.BackColor = Color.SteelBlue
			lblYearCSAT.ForeColor = Color.White
		End If

		If lblYearAway.Text = "100" Then
			lblYearAway.BackColor = Color.IndianRed
			lblYearAway.ForeColor = Color.Black
		End If

	End Sub

	Private Sub cmbEmployee_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEmployee.SelectedIndexChanged
		'Display the metrics for the selected employee upon the user changing the selected index of employees.

		Me.Cursor = Cursors.WaitCursor

		'Clear the label contents before displaying the metrics for a different employee.
		lblJanDuration.Text = ""
		lblJanCSAT.Text = ""
		lblJanAway.Text = ""
		lblJanQual.Text = ""
		lblJanDev.Text = ""
		lblFebDuration.Text = ""
		lblFebCSAT.Text = ""
		lblFebAway.Text = ""
		lblFebQual.Text = ""
		lblFebDev.Text = ""
		lblMarDuration.Text = ""
		lblMarCSAT.Text = ""
		lblMarAway.Text = ""
		lblMarQual.Text = ""
		lblMarDev.Text = ""
		lblAprDuration.Text = ""
		lblAprCSAT.Text = ""
		lblAprAway.Text = ""
		lblAprQual.Text = ""
		lblAprDev.Text = ""
		lblMayDuration.Text = ""
		lblMayCSAT.Text = ""
		lblMayAway.Text = ""
		lblMayQual.Text = ""
		lblMayDev.Text = ""
		lblJunDuration.Text = ""
		lblJunCSAT.Text = ""
		lblJunAway.Text = ""
		lblJunQual.Text = ""
		lblJunDev.Text = ""
		lblJulDuration.Text = ""
		lblJulCSAT.Text = ""
		lblJulAway.Text = ""
		lblJulQual.Text = ""
		lblJulDev.Text = ""
		lblAugDuration.Text = ""
		lblAugCSAT.Text = ""
		lblAugAway.Text = ""
		lblAugQual.Text = ""
		lblAugDev.Text = ""
		lblSepDuration.Text = ""
		lblSepCSAT.Text = ""
		lblSepAway.Text = ""
		lblSepQual.Text = ""
		lblSepDev.Text = ""
		lblOctDuration.Text = ""
		lblOctCSAT.Text = ""
		lblOctAway.Text = ""
		lblOctQual.Text = ""
		lblOctDev.Text = ""
		lblNovDuration.Text = ""
		lblNovCSAT.Text = ""
		lblNovAway.Text = ""
		lblNovQual.Text = ""
		lblNovDev.Text = ""
		lblDecDuration.Text = ""
		lblDecCSAT.Text = ""
		lblDecAway.Text = ""
		lblDecQual.Text = ""
		lblDecDev.Text = ""
		lblYearDuration.Text = ""
		lblYearCSAT.Text = ""
		lblYearAway.Text = ""
		lblYearQual.Text = ""
		lblYearDev.Text = ""

		DisplayJanMetrics()
		DisplayFebMetrics()
		DisplayMarMetrics()
		DisplayAprMetrics()
		DisplayMayMetrics()
		DisplayJunMetrics()
		DisplayJulMetrics()
		DisplayAugMetrics()
		DisplaySepMetrics()
		DisplayOctMetrics()
		DisplayNovMetrics()
		DisplayDecMetrics()
		ColorJanMetrics()
		ColorFebMetrics()
		ColorMarMetrics()
		ColorAprMetrics()
		ColorMayMetrics()
		ColorJunMetrics()
		ColorJulMetrics()
		ColorAugMetrics()
		ColorSepMetrics()
		ColorOctMetrics()
		ColorNovMetrics()
		ColorDecMetrics()

		Me.Cursor = Cursors.Default
	End Sub

	Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
		'Declare variables for the CSAT metric and yearly average. CSAT variables with integers coincide with the statements below depending on how many months have data.
		Dim dblCSAT1 As Double
		Dim dblCSAT2 As Double
		Dim dblCSAT3 As Double
		Dim dblCSAT4 As Double
		Dim dblCSAT5 As Double
		Dim dblCSAT6 As Double
		Dim dblCSAT7 As Double
		Dim dblCSAT8 As Double
		Dim dblCSAT9 As Double
		Dim dblCSAT10 As Double
		Dim dblCSAT11 As Double
		Dim dblCSAT12 As Double
		Dim dblJanCSAT As Double
		Dim dblFebCSAT As Double
		Dim dblMarCSAT As Double
		Dim dblAprCSAT As Double
		Dim dblMayCSAT As Double
		Dim dblJunCSAT As Double
		Dim dblJulCSAT As Double
		Dim dblAugCSAT As Double
		Dim dblSepCSAT As Double
		Dim dblOctCSAT As Double
		Dim dblNovCSAT As Double
		Dim dblDecCSAT As Double

		Double.TryParse(lblJanCSAT.Text, dblJanCSAT)
		Double.TryParse(lblFebCSAT.Text, dblFebCSAT)
		Double.TryParse(lblMarCSAT.Text, dblMarCSAT)
		Double.TryParse(lblAprCSAT.Text, dblAprCSAT)
		Double.TryParse(lblMayCSAT.Text, dblMayCSAT)
		Double.TryParse(lblJunCSAT.Text, dblJunCSAT)
		Double.TryParse(lblJulCSAT.Text, dblJulCSAT)
		Double.TryParse(lblAugCSAT.Text, dblAugCSAT)
		Double.TryParse(lblSepCSAT.Text, dblSepCSAT)
		Double.TryParse(lblOctCSAT.Text, dblOctCSAT)
		Double.TryParse(lblNovCSAT.Text, dblNovCSAT)
		Double.TryParse(lblDecCSAT.Text, dblDecCSAT)

		'Calculate the average for CSAT if all months have metric data, if not get the average by how many months do have metric data.
		If lblJanCSAT.Text IsNot "" Then
			dblCSAT1 = dblJanCSAT
			lblYearCSAT.Text = dblCSAT1.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" Then
			dblCSAT2 = (dblJanCSAT + dblFebCSAT) / 2
			lblYearCSAT.Text = dblCSAT2.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" Then
			dblCSAT3 = (dblJanCSAT + dblFebCSAT + dblMarCSAT) / 3
			lblYearCSAT.Text = dblCSAT3.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" Then
			dblCSAT4 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT) / 4
			lblYearCSAT.Text = dblCSAT4.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" Then
			dblCSAT5 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT) / 5
			lblYearCSAT.Text = dblCSAT5.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" And lblJunCSAT.Text IsNot "" Then
			dblCSAT6 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT + dblJunCSAT) / 6
			lblYearCSAT.Text = dblCSAT6.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" And lblJunCSAT.Text IsNot "" And lblJulCSAT.Text IsNot "" Then
			dblCSAT7 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT + dblJunCSAT + dblJulCSAT) / 7
			lblYearCSAT.Text = dblCSAT7.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" And lblJunCSAT.Text IsNot "" And lblJulCSAT.Text IsNot "" And lblAugCSAT.Text IsNot "" Then
			dblCSAT8 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT + dblJunCSAT + dblJulCSAT + dblAugCSAT) / 8
			lblYearCSAT.Text = dblCSAT8.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" And lblJunCSAT.Text IsNot "" And lblJulCSAT.Text IsNot "" And lblAugCSAT.Text IsNot "" And lblSepCSAT.Text IsNot "" Then
			dblCSAT9 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT + dblJunCSAT + dblJulCSAT + dblAugCSAT + dblSepCSAT) / 9
			lblYearCSAT.Text = dblCSAT9.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" And lblJunCSAT.Text IsNot "" And lblJulCSAT.Text IsNot "" And lblAugCSAT.Text IsNot "" And lblSepCSAT.Text IsNot "" And lblOctCSAT.Text IsNot "" Then
			dblCSAT10 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT + dblJunCSAT + dblJulCSAT + dblAugCSAT + dblSepCSAT + dblOctCSAT) / 10
			lblYearCSAT.Text = dblCSAT10.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" And lblJunCSAT.Text IsNot "" And lblJulCSAT.Text IsNot "" And lblAugCSAT.Text IsNot "" And lblSepCSAT.Text IsNot "" And lblOctCSAT.Text IsNot "" And lblNovCSAT.Text IsNot "" Then
			dblCSAT11 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT + dblJunCSAT + dblJulCSAT + dblAugCSAT + dblSepCSAT + dblOctCSAT + dblNovCSAT) / 11
			lblYearCSAT.Text = dblCSAT11.ToString("n2")
		End If

		If lblJanCSAT.Text IsNot "" And lblFebCSAT.Text IsNot "" And lblMarCSAT.Text IsNot "" And lblAprCSAT.Text IsNot "" And lblMayCSAT.Text IsNot "" And lblJunCSAT.Text IsNot "" And lblJulCSAT.Text IsNot "" And lblAugCSAT.Text IsNot "" And lblSepCSAT.Text IsNot "" And lblOctCSAT.Text IsNot "" And lblNovCSAT.Text IsNot "" And lblDecCSAT.Text IsNot "" Then
			dblCSAT12 = (dblJanCSAT + dblFebCSAT + dblMarCSAT + dblAprCSAT + dblMayCSAT + dblJunCSAT + dblJulCSAT + dblAugCSAT + dblSepCSAT + dblOctCSAT + dblNovCSAT + dblDecCSAT) / 12
			lblYearCSAT.Text = dblCSAT12.ToString("n2")
		End If

		ColorYearMetrics()

	End Sub
End Class