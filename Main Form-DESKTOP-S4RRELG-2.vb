'Allstate Service Delivery Coaching Tool
'Developed by Chris DePaul using VB.NET
'Version 1.0 created on 10/26/2021

Option Strict On
Imports System.Data.OleDb

Public Class frmMain
	'Declare sub procedures to input the monthly metrics.
	Private Sub AddJanMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		Dim query As String = String.Empty
		query &= "INSERT INTO January (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddFebMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO February (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddMarMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO March (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddAprMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO April (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddMayMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO May (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddJunMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO June (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddJulMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO July (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddAugMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO August (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddSepMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO September (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddOctMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO October (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddNovMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO November (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	Private Sub AddDecMetrics()
		'Declare variables
		Dim strManager = cmbManager.SelectedItem.ToString
		Dim strEmployee = cmbEmployee.SelectedItem.ToString
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO December (Manager, Employee, Duration, CSAT, Away, Quality, Development) VALUES (@Manager, @Employee, @Duration, @CSAT, @Away, @Quality, @Development)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Duration", strDuration)
					.Parameters.AddWithValue("@CSAT", strCSAT)
					.Parameters.AddWithValue("@Away", strAway)
					.Parameters.AddWithValue("@Quality", strQuality)
					.Parameters.AddWithValue("@Development", strDevelopment)
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

	'Declare sub procedures to display the monthly metrics.
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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
			Dim sql As String = "Select Duration, CSAT, Away, Quality, Development FROM February WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

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
				lblDuration.Text = myReader("Duration").ToString
				lblCSAT.Text = myReader("CSAT").ToString
				lblAway.Text = myReader("Away").ToString
				lblQuality.Text = myReader("Quality").ToString
				lblDevelopment.Text = myReader("Development").ToString
			End While

			'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
			If lblDuration.Text <= "16:00" Then
				lblDuration.BackColor = Color.SteelBlue
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "16:00" Then
				lblDuration.BackColor = Color.ForestGreen
				lblDuration.ForeColor = Color.White
			End If

			If lblDuration.Text > "21:00" Then
				lblDuration.BackColor = Color.Gold
				lblDuration.ForeColor = Color.Black
			End If

			If lblDuration.Text > "25:00" Then
				lblDuration.BackColor = Color.IndianRed
				lblDuration.ForeColor = Color.Black
			End If

			If lblCSAT.Text >= "87.00" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "86.99" Then
				lblCSAT.BackColor = Color.ForestGreen
				lblCSAT.ForeColor = Color.White
			End If

			If lblCSAT.Text <= "74.99" Then
				lblCSAT.BackColor = Color.Gold
				lblCSAT.ForeColor = Color.Black
			End If

			If lblCSAT.Text < "70.00" Then
				lblCSAT.BackColor = Color.IndianRed
				lblCSAT.ForeColor = Color.Black
			End If

			If lblAway.Text <= "26.00" Then
				lblAway.BackColor = Color.SteelBlue
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "26.00" Then
				lblAway.BackColor = Color.ForestGreen
				lblAway.ForeColor = Color.White
			End If

			If lblAway.Text > "29.99" Then
				lblAway.BackColor = Color.Gold
				lblAway.ForeColor = Color.Black
			End If

			If lblAway.Text > "37.00" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If

			If lblDuration.Text Is "" Then
				lblDuration.BackColor = Color.White
			End If

			If lblCSAT.Text Is "" Then
				lblCSAT.BackColor = Color.White
			End If

			If lblAway.Text Is "" Then
				lblAway.BackColor = Color.White
			End If

			If lblQuality.Text Is "" Then
				lblQuality.BackColor = Color.White
			End If

			If lblDevelopment.Text Is "" Then
				lblDevelopment.BackColor = Color.White
			End If

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	'Add/Update/Display/Delete additional feedback for the current employee.
	Private Sub AddFeedback()
		'Declare variables
		Dim strManager = CStr(cmbManager.SelectedItem.ToString)
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strFeedback = CStr(txtAddFeedback.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO AddFeedback (Manager, Employee, AddFeedback) VALUES (@Manager, @Employee, @AddFeedback)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@AddFeedback", strFeedback)
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

	Private Sub UpdateAddFeedback()
		'Declare variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strFeedback = CStr(txtAddFeedback.Text)
		Dim query As String = String.Empty
		query &= "UPDATE AddFeedback SET AddFeedback = @AddFeedback WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@AddFeedback", strFeedback)
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

	Private Sub DisplayAddFeedback()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT AddFeedback FROM AddFeedback WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				txtAddFeedback.Text = myReader("AddFeedback").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DeleteAddFeedback()
		Dim query As String = String.Empty
		query &= "DELETE FROM AddFeedback WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	'Add/Update/Display/Delete general notes for the current employee.
	Private Sub AddGeneralNotes()
		'Declare variables
		Dim strManager = CStr(cmbManager.SelectedItem.ToString)
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strGeneral = CStr(txtGeneralNotes.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO GeneralNotes (Manager, Employee, GeneralNotes) VALUES (@Manager, @Employee, @General)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@General", strGeneral)
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

	Private Sub UpdateGeneralNotes()
		'Declare variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strGeneral = CStr(txtGeneralNotes.Text)
		Dim query As String = String.Empty
		query &= "UPDATE GeneralNotes SET GeneralNotes = @General WHERE Employee ='" & cmbEmployee.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@General", strGeneral)
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

	Private Sub DisplayGeneralNotes()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT GeneralNotes FROM GeneralNotes WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				txtGeneralNotes.Text = myReader("GeneralNotes").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	Private Sub DeleteGeneralNotes()
		Dim query As String = String.Empty
		query &= "DELETE FROM GeneralNotes WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	'Add the behaviors for the current employee.
	Private Sub AddBehaviors()
		'Declare variables
		Dim strManager = CStr(cmbManager.SelectedItem.ToString)
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strChallenge = CStr(txtChallenge.Text)
		Dim strCollaborate = CStr(txtCollaborate.Text)
		Dim strFeedback = CStr(txtFeedback.Text)
		Dim strClarity = CStr(txtClarity1.Text)
		Dim query As String = String.Empty
		query &= "INSERT INTO Behaviors (Manager, Employee, Challenge, Collaborate, Feedback, Clarity) VALUES (@Manager, @Employee, @Challenge, @Collaborate, @Feedback, @Clarity)"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Manager", strManager)
					.Parameters.AddWithValue("@Employee", strEmployee)
					.Parameters.AddWithValue("@Challenge", strChallenge)
					.Parameters.AddWithValue("@Collaborate", strCollaborate)
					.Parameters.AddWithValue("@Feedback", strFeedback)
					.Parameters.AddWithValue("@Clarity", strClarity)
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

	'Display the behaviors for the current employee.
	Private Sub DisplayBehaviors()
		'Open a connection to the database and then assign the values from the appropriate metric columns to the appropriate labels.
		Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb"
		Dim conn As New OleDbConnection(str)

		Try
			conn.Open()
			Dim sql As String = "SELECT Challenge, Collaborate, Feedback, Clarity FROM Behaviors WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"
			Dim cmd As New OleDbCommand(sql, conn)
			Dim myReader As OleDbDataReader = cmd.ExecuteReader()

			While myReader.Read
				txtChallenge.Text = myReader("Challenge").ToString
				txtCollaborate.Text = myReader("Collaborate").ToString
				txtFeedback.Text = myReader("Feedback").ToString
				txtClarity1.Text = myReader("Clarity").ToString
			End While

		Catch ex As OleDbException
			MsgBox(ex.ToString)
		Finally
			conn.Close()
		End Try
	End Sub

	'Update the current behaviors if there is already behavior comments in the database.
	Private Sub UpdateBehaviors()
		'Declare variables.
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strChallenge = CStr(txtChallenge.Text)
		Dim strCollaborate = CStr(txtCollaborate.Text)
		Dim strFeedback = CStr(txtFeedback.Text)
		Dim strClarity = CStr(txtClarity1.Text)
		Dim query As String = String.Empty
		query &= "UPDATE Behaviors SET Challenge = @Challenge, Collaborate = @Collaborate, Feedback = @Feedback, Clarity = @Clarity WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query
					.Parameters.AddWithValue("@Challenge", strChallenge)
					.Parameters.AddWithValue("@Collaborate", strCollaborate)
					.Parameters.AddWithValue("@Feedback", strFeedback)
					.Parameters.AddWithValue("@Clarity", strClarity)
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

	'Update the current metrics for the employees in case there were only partial results originally entered.
	Private Sub UpdateJanMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth = "January" Then
			Dim query As String = String.Empty
			query &= "UPDATE January SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth = "January" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE January SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth = "January" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE January SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth = "January" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE January SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth = "January" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE January SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If

	End Sub

	Private Sub UpdateFebMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth = "February" Then
			Dim query As String = String.Empty
			query &= "UPDATE February SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth = "February" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE February SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth = "February" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE February SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth = "February" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE February SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth = "February" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE February SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateMarMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "March" Then
			Dim query As String = String.Empty
			query &= "UPDATE March SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "March" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE March SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "March" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE March SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "March" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE March SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "March" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE March SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateAprMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "April" Then
			Dim query As String = String.Empty
			query &= "UPDATE April SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "April" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE April SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "April" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE April SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "April" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE April SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "April" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE April SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateMayMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "May" Then
			Dim query As String = String.Empty
			query &= "UPDATE May SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "May" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE May SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "May" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE May SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "May" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE May SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "May" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE May SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateJunMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "June" Then
			Dim query As String = String.Empty
			query &= "UPDATE June SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "June" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE June SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "June" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE June SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "June" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE June SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "June" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE June SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateJulMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "July" Then
			Dim query As String = String.Empty
			query &= "UPDATE July SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "July" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE July SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "July" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE July SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "July" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE July SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "July" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE July SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateAugMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "August" Then
			Dim query As String = String.Empty
			query &= "UPDATE August SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "August" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE August SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "August" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE August SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "August" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE August SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "August" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE August SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateSepMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "September" Then
			Dim query As String = String.Empty
			query &= "UPDATE September SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "September" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE September SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "September" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE September SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "September" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE September SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "September" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE September SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateOctMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "October" Then
			Dim query As String = String.Empty
			query &= "UPDATE October SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "October" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE October SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "October" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE October SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "October" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE October SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "October" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE October SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateNovMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "November" Then
			Dim query As String = String.Empty
			query &= "UPDATE November SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "November" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE November SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "November" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE November SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "November" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE November SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "November" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE November SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	Private Sub UpdateDecMetrics()
		'Declare double variables
		Dim strEmployee = CStr(cmbEmployee.SelectedItem.ToString)
		Dim strMonth = CStr(cmbMonth.SelectedItem.ToString)
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDevelopment = CStr(txtDev.Text)

		If lblDuration.Text Is "" AndAlso strMonth Is "December" Then
			Dim query As String = String.Empty
			query &= "UPDATE December SET Duration = @Duration WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query
						.Parameters.AddWithValue("@Duration", strDuration)
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
		End If

		If lblCSAT.Text Is "" AndAlso strMonth Is "December" Then
			Dim query1 As String = String.Empty
			query1 &= "UPDATE December SET CSAT = @CSAT WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query1
						.Parameters.AddWithValue("@CSAT", strCSAT)
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
		End If

		If lblAway.Text Is "" AndAlso strMonth Is "December" Then
			Dim query2 As String = String.Empty
			query2 &= "UPDATE December SET Away = @Away WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query2
						.Parameters.AddWithValue("@Away", strAway)
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
		End If

		If lblQuality.Text Is "" AndAlso strMonth Is "December" Then
			Dim query3 As String = String.Empty
			query3 &= "UPDATE December SET Quality = @Quality WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query3
						.Parameters.AddWithValue("@Quality", strQuality)
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
		End If

		If lblDevelopment.Text Is "" AndAlso strMonth Is "December" Then
			Dim query4 As String = String.Empty
			query4 &= "UPDATE December SET Development = @Development WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

			Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
				Using comm As New OleDbCommand()
					With comm
						.Connection = conn
						.CommandType = CommandType.Text
						.CommandText = query4
						.Parameters.AddWithValue("@Development", strDevelopment)
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
		End If
	End Sub

	'Delete the metrics or behaviors for the current employee based off the monthly selection from the menustrip.
	Private Sub DeleteJanMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM January WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteFebMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM February WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteMarMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM March WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteAprMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM April WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteMayMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM May WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteJunMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM June WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteJulMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM July WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteAugMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM August WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteSepMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM September WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteOctMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM October WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteNovMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM November WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteDecMetrics()
		Dim query As String = String.Empty
		query &= "DELETE FROM December WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteBehaviors()
		Dim query As String = String.Empty
		query &= "DELETE FROM Behaviors WHERE Employee =" & "'" & cmbEmployee.SelectedItem.ToString & "'"

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

	Private Sub DeleteYear()
		Dim query As String = String.Empty
		query &= "DELETE FROM January WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

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

		Dim query1 As String = String.Empty
		query1 &= "DELETE FROM February WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query1
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

		Dim query2 As String = String.Empty
		query2 &= "DELETE FROM March WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query2
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

		Dim query3 As String = String.Empty
		query3 &= "DELETE FROM April WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query3
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

		Dim query4 As String = String.Empty
		query4 &= "DELETE FROM May WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query4
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

		Dim query5 As String = String.Empty
		query5 &= "DELETE FROM June WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query5
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

		Dim query6 As String = String.Empty
		query6 &= "DELETE FROM July WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query6
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

		Dim query7 As String = String.Empty
		query7 &= "DELETE FROM August WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query7
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

		Dim query8 As String = String.Empty
		query8 &= "DELETE FROM September WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query8
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

		Dim query9 As String = String.Empty
		query9 &= "DELETE FROM October WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query9
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

		Dim query10 As String = String.Empty
		query10 &= "DELETE FROM November WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query10
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

		Dim query11 As String = String.Empty
		query11 &= "DELETE FROM December WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query11
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

		Dim query12 As String = String.Empty
		query12 &= "DELETE FROM Behaviors WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query12
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

		Dim query13 As String = String.Empty
		query13 &= "DELETE FROM GeneralNotes WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query13
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

		Dim query14 As String = String.Empty
		query14 &= "DELETE FROM AddFeedback WHERE Manager='" & cmbManager.SelectedItem.ToString & "'"

		Using conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\CoachingDB.accdb")
			Using comm As New OleDbCommand()
				With comm
					.Connection = conn
					.CommandType = CommandType.Text
					.CommandText = query14
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

	Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

		cmbMonth.Items.Add("January")
		cmbMonth.Items.Add("February")
		cmbMonth.Items.Add("March")
		cmbMonth.Items.Add("April")
		cmbMonth.Items.Add("May")
		cmbMonth.Items.Add("June")
		cmbMonth.Items.Add("July")
		cmbMonth.Items.Add("August")
		cmbMonth.Items.Add("September")
		cmbMonth.Items.Add("October")
		cmbMonth.Items.Add("November")
		cmbMonth.Items.Add("December")

	End Sub

	Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
		'Close the application upon clicking the "Exit" button.
		Me.Close()
	End Sub

	Private Sub mnuViewDap_Click(sender As Object, e As EventArgs) Handles mnuViewDAP.Click
		'Open the Direct Analytics Portal main page

		'Display a wait cursor so that the user knows the webpage is loading and could potentially be in the background.
		Me.Cursor = Cursors.WaitCursor

		'Declare a variable for opening the TalentConnection webpage and then use the variable to load the page in the default web browser.
		Dim strDAP As String = "http://dap"
		Process.Start(strDAP)

		'Return the cursor to the default cursor after the procedure has completed.
		Me.Cursor = Cursors.Default
	End Sub

	Private Sub mnuViewTc_Click(sender As Object, e As EventArgs) Handles mnuViewTC.Click
		'Open TalentConnection main page.

		'Display a wait cursor so that the user knows the webpage is loading and could potentially be in the background.
		Me.Cursor = Cursors.WaitCursor

		'Declare a variable for opening the TalentConnection webpage and then use the variable to load the page in the default web browser.
		Dim strTC As String = "https://agtacc.allstate.com/FIM/sps/successfactors_ag/saml20/logininitial?RequestBinding=HTTPPost&NameIdFormat=email&PartnerId=https://www.successfactors.com"

		Process.Start(strTC)

		'Return the cursor to the default cursor after the procedure has completed.
		Me.Cursor = Cursors.Default
	End Sub

	Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles picAllstate.Click
		'Open the Allstate/AllConnect home page upon clicking the Allstate logo.

		'Display a wait cursor so that the user knows the webpage is loading and could potentially be in the background.
		Me.Cursor = Cursors.WaitCursor

		'Declare a variable for opening the Allstate/AllConnect webpage and then use the variable to load the page in the default web browser.
		Dim strAllstate As String = "https://allstatecloud.sharepoint.com/sites/acn-portal"
		Process.Start(strAllstate)

		'Return the cursor to the default cursor after the procedure has completed.
		Me.Cursor = Cursors.Default
	End Sub

	Private Sub picAllstate_MouseHover(sender As Object, e As EventArgs) Handles picAllstate.MouseHover
		'Change the mouse icon to a HAND when the user hovers over the Allstate Picture Box.
		Me.Cursor = Cursors.Hand
	End Sub

	Private Sub picAllstate_MouseLeave(sender As Object, e As EventArgs) Handles picAllstate.MouseLeave
		'Change the mouse icon back to the default mouse cursor when the user stops hovering over the Allstate Picture Box.
		Me.Cursor = Cursors.Default
	End Sub

	Private Sub mnuHelp_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
		'Show the help form upon clicking the Help button.
		frmHelp.Show()
	End Sub

	Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
		'Clear the textbox input upon clicking the Clear button and also clear the Employee Selection and Month Selection combo boxes.

		'Clear text boxes.
		txtDuration.Clear()
		txtCSAT.Clear()
		txtAway.Clear()
		txtQual.Clear()
		txtDev.Clear()

	End Sub

	Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
		'Determine the month and employee from the selection combo boxes and save the information in the appropriate database table(month)

		'Delcare variables used to determine which employee and month have been selected and if there is already comments for the behaviors.
		Dim strEmployee = CStr(cmbEmployee.SelectedItem)
		Dim strMonth = CStr(cmbMonth.SelectedItem)

		'Change the mouse cursor so that the user knows there are procedures running in the background upon clicking the "Save" button.
		Me.Cursor = Cursors.WaitCursor

		'Use If Then and Select Case to determine which month is selected and insert with the correct employee name and metrics while either inputting or updating behaviors.
		If strMonth Is "January" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddJanMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayJanMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "January" Then
			UpdateJanMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayJanMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "February" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddFebMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayFebMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "February" Then
			UpdateFebMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayFebMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "March" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddMarMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayMarMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "March" Then
			UpdateMarMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayMarMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "April" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddAprMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayAprMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "April" Then
			UpdateAprMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayAprMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "May" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddMayMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayMayMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "May" Then
			UpdateMayMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayMayMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "June" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddJunMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayJunMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "June" Then
			UpdateJunMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayJunMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "July" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddJulMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayJulMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "July" Then
			UpdateJulMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayJulMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "August" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddAugMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayAugMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "August" Then
			UpdateAugMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayAugMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "September" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddSepMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplaySepMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "September" Then
			UpdateSepMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplaySepMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "October" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddOctMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayOctMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "October" Then
			UpdateOctMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayOctMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "November" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddNovMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayNovMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "November" Then
			UpdateNovMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayNovMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		If strMonth Is "December" And lblDuration.Text Is "" And lblCSAT.Text Is "" And lblAway.Text Is "" And lblQuality.Text Is "" And lblDevelopment.Text Is "" Then
			Select Case strEmployee
				Case strEmployee.ToString
					AddDecMetrics()
					AddBehaviors()
					AddFeedback()
					AddGeneralNotes()
					DisplayDecMetrics()

					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If
					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					MessageBox.Show("Metrics/Behaviors Recorded")
			End Select
		ElseIf strMonth Is "December" Then
			UpdateDecMetrics()
			UpdateBehaviors()
			UpdateAddFeedback()
			UpdateGeneralNotes()
			DisplayDecMetrics()
			DisplayBehaviors()
			DisplayAddFeedback()
			DisplayGeneralNotes()

			If lblCSAT.Text = "100" Then
				lblCSAT.BackColor = Color.SteelBlue
				lblCSAT.ForeColor = Color.White
			End If
			If lblAway.Text = "100" Then
				lblAway.BackColor = Color.IndianRed
				lblAway.ForeColor = Color.Black
			End If
			MessageBox.Show("Metrics/Behaviors Updated")
		End If

		'Clear all textboxes when the user clicks the Save button.
		txtDuration.Clear()
		txtCSAT.Clear()
		txtAway.Clear()
		txtQual.Clear()
		txtDev.Clear()

		'Return the cursor to default after the procedure has fully run.
		Me.Cursor = Cursors.Default

		If lblCSAT.Text = "100" Then
			lblCSAT.BackColor = Color.SteelBlue
			lblCSAT.ForeColor = Color.White
		End If

		If lblAway.Text = "100" Then
			lblAway.BackColor = Color.IndianRed
			lblAway.ForeColor = Color.Black
		End If

		'If no information has been provided keep the labels as the default white background color.
		If lblDuration.Text Is "" Then
			lblDuration.BackColor = Color.White
		End If

		If lblCSAT.Text Is "" Then
			lblCSAT.BackColor = Color.White
		End If

		If lblAway.Text Is "" Then
			lblAway.BackColor = Color.White
		End If

		If lblQuality.Text Is "" Then
			lblQuality.BackColor = Color.White
		End If

		If lblDevelopment.Text Is "" Then
			lblDevelopment.BackColor = Color.White
		End If

	End Sub

	Private Sub cmbEmployee_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEmployee.SelectedIndexChanged
		'Declare variables to be used in the below if then/select case statements.
		Dim strEmployee = CStr(cmbEmployee.SelectedItem)
		Dim strMonth = CStr(cmbMonth.SelectedItem)

		txtChallenge.Clear()
		txtCollaborate.Clear()
		txtFeedback.Clear()
		txtClarity1.Clear()
		txtAddFeedback.Clear()
		txtGeneralNotes.Clear()

		'Display the appropriate metric based off employee and month selection.
		If strMonth Is "January" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayJanMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "February" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayFebMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "March" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayMarMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "April" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayAprMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "May" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayMayMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "June" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayJunMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "July" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayJulMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "August" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayAugMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "September" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplaySepMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "October" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayOctMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "November" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayNovMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "December" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayDecMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblDuration.Text <= "16:00" Then
			lblDuration.BackColor = Color.SteelBlue
			lblDuration.ForeColor = Color.White
		End If

		If lblDuration.Text > "16:00" Then
			lblDuration.BackColor = Color.ForestGreen
			lblDuration.ForeColor = Color.White
		End If

		If lblDuration.Text > "21:00" Then
			lblDuration.BackColor = Color.Gold
			lblDuration.ForeColor = Color.Black
		End If

		If lblDuration.Text > "25:00" Then
			lblDuration.BackColor = Color.IndianRed
			lblDuration.ForeColor = Color.Black
		End If

		If lblCSAT.Text >= "87.00" Then
			lblCSAT.BackColor = Color.SteelBlue
			lblCSAT.ForeColor = Color.White
		End If

		If lblCSAT.Text <= "86.99" Then
			lblCSAT.BackColor = Color.ForestGreen
			lblCSAT.ForeColor = Color.White
		End If

		If lblCSAT.Text <= "74.99" Then
			lblCSAT.BackColor = Color.Gold
			lblCSAT.ForeColor = Color.Black
		End If

		If lblCSAT.Text < "70.00" Then
			lblCSAT.BackColor = Color.IndianRed
			lblCSAT.ForeColor = Color.Black
		End If

		If lblAway.Text <= "26.00" Then
			lblAway.BackColor = Color.SteelBlue
			lblAway.ForeColor = Color.White
		End If

		If lblAway.Text > "26.00" Then
			lblAway.BackColor = Color.ForestGreen
			lblAway.ForeColor = Color.White
		End If

		If lblAway.Text > "29.99" Then
			lblAway.BackColor = Color.Gold
			lblAway.ForeColor = Color.Black
		End If

		If lblAway.Text > "37.00" Then
			lblAway.BackColor = Color.IndianRed
			lblAway.ForeColor = Color.Black
		End If

		If lblCSAT.Text = "100" Then
			lblCSAT.BackColor = Color.SteelBlue
			lblCSAT.ForeColor = Color.White
		End If

		If lblAway.Text = "100" Then
			lblAway.BackColor = Color.IndianRed
			lblAway.ForeColor = Color.Black
		End If

		If lblDuration.Text Is "" Then
			lblDuration.BackColor = Color.White
		End If

		If lblCSAT.Text Is "" Then
			lblCSAT.BackColor = Color.White
		End If

		If lblAway.Text Is "" Then
			lblAway.BackColor = Color.White
		End If

		If lblQuality.Text Is "" Then
			lblQuality.BackColor = Color.White
		End If

		If lblDevelopment.Text Is "" Then
			lblDevelopment.BackColor = Color.White
		End If

	End Sub

	Private Sub cmbMonth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMonth.SelectedIndexChanged
		'Declare variables to be used in the below if then/select case statements.
		Dim strEmployee = CStr(cmbEmployee.SelectedItem)
		Dim strMonth = CStr(cmbMonth.SelectedItem)

		'Display the appropriate metric based off employee and month selection.
		If strMonth Is "January" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayJanMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "February" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayFebMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "March" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayMarMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "April" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayAprMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "May" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayMayMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "June" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayJunMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "July" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayJulMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "August" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayAugMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "September" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplaySepMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "October" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayOctMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "November" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayNovMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		If strMonth Is "December" Then
			Select Case strEmployee
				Case strEmployee.ToString
					DisplayDecMetrics()
					If lblCSAT.Text = "100" Then
						lblCSAT.BackColor = Color.SteelBlue
						lblCSAT.ForeColor = Color.White
					End If

					If lblAway.Text = "100" Then
						lblAway.BackColor = Color.IndianRed
						lblAway.ForeColor = Color.Black
					End If
					DisplayBehaviors()
					DisplayAddFeedback()
					DisplayGeneralNotes()
			End Select
		End If

		'Display the appropriate color for BTE, Expected, Inconsistent, or Unacceptable for the employees metrics.
		If lblDuration.Text <= "16:00" Then
			lblDuration.BackColor = Color.SteelBlue
			lblDuration.ForeColor = Color.White
		End If

		If lblDuration.Text > "16:00" Then
			lblDuration.BackColor = Color.ForestGreen
			lblDuration.ForeColor = Color.White
		End If

		If lblDuration.Text > "21:00" Then
			lblDuration.BackColor = Color.Gold
			lblDuration.ForeColor = Color.Black
		End If

		If lblDuration.Text > "25:00" Then
			lblDuration.BackColor = Color.IndianRed
			lblDuration.ForeColor = Color.Black
		End If

		If lblCSAT.Text >= "87.00" Then
			lblCSAT.BackColor = Color.SteelBlue
			lblCSAT.ForeColor = Color.White
		End If

		If lblCSAT.Text <= "86.99" Then
			lblCSAT.BackColor = Color.ForestGreen
			lblCSAT.ForeColor = Color.White
		End If

		If lblCSAT.Text <= "74.99" Then
			lblCSAT.BackColor = Color.Gold
			lblCSAT.ForeColor = Color.Black
		End If

		If lblCSAT.Text < "70.00" Then
			lblCSAT.BackColor = Color.IndianRed
			lblCSAT.ForeColor = Color.Black
		End If

		If lblAway.Text <= "26.00" Then
			lblAway.BackColor = Color.SteelBlue
			lblAway.ForeColor = Color.White
		End If

		If lblAway.Text > "26.00" Then
			lblAway.BackColor = Color.ForestGreen
			lblAway.ForeColor = Color.White
		End If

		If lblAway.Text > "29.99" Then
			lblAway.BackColor = Color.Gold
			lblAway.ForeColor = Color.Black
		End If

		If lblAway.Text > "37.00" Then
			lblAway.BackColor = Color.IndianRed
			lblAway.ForeColor = Color.Black
		End If

		If lblCSAT.Text = "100" Then
			lblCSAT.BackColor = Color.SteelBlue
			lblCSAT.ForeColor = Color.White
		End If

		If lblAway.Text = "100" Then
			lblAway.BackColor = Color.IndianRed
			lblAway.ForeColor = Color.Black
		End If

		If lblDuration.Text Is "" Then
			lblDuration.BackColor = Color.White
		End If

		If lblCSAT.Text Is "" Then
			lblCSAT.BackColor = Color.White
		End If

		If lblAway.Text Is "" Then
			lblAway.BackColor = Color.White
		End If

		If lblQuality.Text Is "" Then
			lblQuality.BackColor = Color.White
		End If

		If lblDevelopment.Text Is "" Then
			lblDevelopment.BackColor = Color.White
		End If

	End Sub

	Private Sub cmbManager_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbManager.SelectedIndexChanged
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

		txtChallenge.Clear()
		txtCollaborate.Clear()
		txtFeedback.Clear()
		txtClarity1.Clear()
		txtAddFeedback.Clear()
		txtGeneralNotes.Clear()

	End Sub

	Private Sub txtDuration_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDuration.KeyPress, txtCSAT.KeyPress, txtAway.KeyPress, txtQual.KeyPress, txtDev.KeyPress
		'Declare the variables used in the below if statements to determine if the user can continue to input digits.
		Dim strDuration = CStr(txtDuration.Text)
		Dim strCSAT = CStr(txtCSAT.Text)
		Dim strAway = CStr(txtAway.Text)
		Dim strQuality = CStr(txtQual.Text)
		Dim strDev = CStr(txtDev.Text)

		'Ensure only numbers,  semicolon, or periods are used in the metric input.
		If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso e.KeyChar <> ControlChars.Back AndAlso e.KeyChar <> ":" AndAlso e.KeyChar <> "." Then
			e.Handled = True
		End If

		'Ensure the user does not improperly log any metrics with additional digits.
		If strDuration.Length > 6 Then
			e.Handled = True
		End If

		If strCSAT.Length > 6 Then
			e.Handled = True
		End If

		If strAway.Length > 6 Then
			e.Handled = True
		End If

		If strQuality.Length > 6 Then
			e.Handled = True
		End If

		If strDev.Length > 6 Then
			e.Handled = True
		End If
	End Sub

	'Allow the user to delete the current metrics for the selected employee and month.
	Private Sub JanuaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsJan.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for January?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteJanMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub FebruaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsFeb.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for February?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteFebMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub MarchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsMar.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for March?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteMarMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub AprilToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsApr.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for April?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteAprMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub MayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsMay.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for May?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteMayMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub JuneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsJun.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for June?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteJunMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub JulyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsJul.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for July?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteJulMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub AugustToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsAug.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for August?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteAugMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub SeptemberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsSep.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for September?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteSepMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub OctoberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsOct.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for October?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteOctMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub NovemberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsNov.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for November?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteNovMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	Private Sub DecemberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteMetricsDec.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the current employees metrics for December?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteDecMetrics()
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
		End If
	End Sub

	'Delete all database entries for metrics and behaviors.
	Private Sub mnuFileDeleteYear_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteYear.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete ALL metrics and ALL notes for your employees for the year? This action cannot be undone.", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		Me.Cursor = Cursors.WaitCursor

		If dlgResult = DialogResult.Yes Then
			DeleteYear()
			MessageBox.Show("All metrics and all notes have been deleted.")
			lblDuration.Text = ""
			lblDuration.BackColor = Color.White
			lblCSAT.Text = ""
			lblCSAT.BackColor = Color.White
			lblAway.Text = ""
			lblAway.BackColor = Color.White
			lblQuality.Text = ""
			lblQuality.BackColor = Color.White
			lblDevelopment.Text = ""
			lblDevelopment.BackColor = Color.White
			txtChallenge.Clear()
			txtCollaborate.Clear()
			txtFeedback.Clear()
			txtClarity1.Clear()
			txtAddFeedback.Clear()
			txtGeneralNotes.Clear()
		End If

		Me.Cursor = Cursors.Default

	End Sub

	'Delete the behaviors for the currently selected employee.
	Private Sub mnuFileDeleteBehaviors_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteBehaviors.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the behaviors for the current employee? This action cannot be undone.", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteBehaviors()
			'Clear the behavior textboxes upon deletion.
			txtChallenge.Clear()
			txtCollaborate.Clear()
			txtFeedback.Clear()
			txtClarity1.Clear()
			MessageBox.Show("All behaviors have been deleted.")
		End If

	End Sub

	'Clear the labels for metrics - if there is data in the database for the employee, their stats will be displayed via the display subprocedures.
	Private Sub cmbEmployee_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbEmployee.SelectedValueChanged
		'Clear the labels so that the metric input unlocks when switching between employees.
		lblDuration.Text = ""
		lblCSAT.Text = ""
		lblAway.Text = ""
		lblQuality.Text = ""
		lblDevelopment.Text = ""

	End Sub

	Private Sub cmbMonth_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbMonth.SelectedValueChanged
		'Clear the labels so that the metric input unlocks when switching between employees.
		lblDuration.Text = ""
		lblCSAT.Text = ""
		lblAway.Text = ""
		lblQuality.Text = ""
		lblDevelopment.Text = ""

	End Sub

	'View the form that has the behavior templates to make it easier to write behaviors.
	Private Sub mnuViewBehaviors_Click(sender As Object, e As EventArgs) Handles mnuViewBehaviors.Click
		frmBehaviors.Show()
	End Sub

	Private Sub GeneralNotesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteGeneralNotes.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the general notes for the current employee? This action cannot be undone.", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteGeneralNotes()
			'Clear the behavior textboxes upon deletion.
			txtGeneralNotes.Clear()
			MessageBox.Show("All general notes have been deleted.")
		End If
	End Sub

	Private Sub AdditionalFeedbackToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuFileDeleteAddFeedback.Click
		Dim dlgResult As DialogResult = MessageBox.Show("Delete the additional feedback for the current employee? This action cannot be undone.", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

		If dlgResult = DialogResult.Yes Then
			DeleteAddFeedback()
			'Clear the behavior textboxes upon deletion.
			txtAddFeedback.Clear()
			MessageBox.Show("All additional feedback notes have been deleted.")
		End If
	End Sub

	Private Sub mnuUpdate_Click(sender As Object, e As EventArgs) Handles mnuAdminLE.Click
		frmUpdate.Show()
	End Sub
End Class