<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdate
	Inherits System.Windows.Forms.Form

	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdate))
		Me.Label1 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Label4 = New System.Windows.Forms.Label()
		Me.Label5 = New System.Windows.Forms.Label()
		Me.Label6 = New System.Windows.Forms.Label()
		Me.Label7 = New System.Windows.Forms.Label()
		Me.Label8 = New System.Windows.Forms.Label()
		Me.cmbCurrentManager = New System.Windows.Forms.ComboBox()
		Me.cmbRemoveManager = New System.Windows.Forms.ComboBox()
		Me.txtAddManager = New System.Windows.Forms.TextBox()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.btnRemoveManager = New System.Windows.Forms.Button()
		Me.btnAddManager = New System.Windows.Forms.Button()
		Me.GroupBox2 = New System.Windows.Forms.GroupBox()
		Me.btnRemoveEmp = New System.Windows.Forms.Button()
		Me.btnTransferEmp = New System.Windows.Forms.Button()
		Me.cmbTransferManager = New System.Windows.Forms.ComboBox()
		Me.cmbToManager = New System.Windows.Forms.ComboBox()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.btnAddEmployee = New System.Windows.Forms.Button()
		Me.cmbRemoveEmployee = New System.Windows.Forms.ComboBox()
		Me.txtAddEmployee = New System.Windows.Forms.TextBox()
		Me.cmbCurrentEmployee = New System.Windows.Forms.ComboBox()
		Me.TextBox1 = New System.Windows.Forms.TextBox()
		Me.GroupBox1.SuspendLayout()
		Me.GroupBox2.SuspendLayout()
		Me.SuspendLayout()
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(3, 28)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(57, 15)
		Me.Label1.TabIndex = 3
		Me.Label1.Text = "Manager:"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(6, 28)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(62, 15)
		Me.Label2.TabIndex = 4
		Me.Label2.Text = "Employee:"
		'
		'Label4
		'
		Me.Label4.AutoSize = True
		Me.Label4.Location = New System.Drawing.Point(3, 152)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(108, 15)
		Me.Label4.TabIndex = 6
		Me.Label4.Text = "Remove Employee:"
		'
		'Label5
		'
		Me.Label5.AutoSize = True
		Me.Label5.Location = New System.Drawing.Point(3, 152)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(103, 15)
		Me.Label5.TabIndex = 7
		Me.Label5.Text = "Remove Manager:"
		'
		'Label6
		'
		Me.Label6.AutoSize = True
		Me.Label6.Location = New System.Drawing.Point(195, 28)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(100, 15)
		Me.Label6.TabIndex = 8
		Me.Label6.Text = "Current Manager:"
		'
		'Label7
		'
		Me.Label7.AutoSize = True
		Me.Label7.Location = New System.Drawing.Point(3, 90)
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(82, 15)
		Me.Label7.TabIndex = 9
		Me.Label7.Text = "Add Manager:"
		'
		'Label8
		'
		Me.Label8.AutoSize = True
		Me.Label8.Location = New System.Drawing.Point(6, 90)
		Me.Label8.Name = "Label8"
		Me.Label8.Size = New System.Drawing.Size(87, 15)
		Me.Label8.TabIndex = 10
		Me.Label8.Text = "Add Employee:"
		'
		'cmbCurrentManager
		'
		Me.cmbCurrentManager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbCurrentManager.FormattingEnabled = True
		Me.cmbCurrentManager.Location = New System.Drawing.Point(6, 46)
		Me.cmbCurrentManager.Name = "cmbCurrentManager"
		Me.cmbCurrentManager.Size = New System.Drawing.Size(162, 23)
		Me.cmbCurrentManager.Sorted = True
		Me.cmbCurrentManager.TabIndex = 0
		'
		'cmbRemoveManager
		'
		Me.cmbRemoveManager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbRemoveManager.FormattingEnabled = True
		Me.cmbRemoveManager.Location = New System.Drawing.Point(6, 170)
		Me.cmbRemoveManager.Name = "cmbRemoveManager"
		Me.cmbRemoveManager.Size = New System.Drawing.Size(162, 23)
		Me.cmbRemoveManager.Sorted = True
		Me.cmbRemoveManager.TabIndex = 2
		'
		'txtAddManager
		'
		Me.txtAddManager.Location = New System.Drawing.Point(6, 108)
		Me.txtAddManager.Name = "txtAddManager"
		Me.txtAddManager.Size = New System.Drawing.Size(162, 23)
		Me.txtAddManager.TabIndex = 1
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.btnRemoveManager)
		Me.GroupBox1.Controls.Add(Me.btnAddManager)
		Me.GroupBox1.Controls.Add(Me.txtAddManager)
		Me.GroupBox1.Controls.Add(Me.Label1)
		Me.GroupBox1.Controls.Add(Me.cmbRemoveManager)
		Me.GroupBox1.Controls.Add(Me.Label5)
		Me.GroupBox1.Controls.Add(Me.cmbCurrentManager)
		Me.GroupBox1.Controls.Add(Me.Label7)
		Me.GroupBox1.Location = New System.Drawing.Point(12, 256)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(175, 268)
		Me.GroupBox1.TabIndex = 14
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Admin: Manager"
		'
		'btnRemoveManager
		'
		Me.btnRemoveManager.FlatAppearance.BorderColor = System.Drawing.Color.SteelBlue
		Me.btnRemoveManager.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
		Me.btnRemoveManager.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SteelBlue
		Me.btnRemoveManager.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.btnRemoveManager.Location = New System.Drawing.Point(69, 213)
		Me.btnRemoveManager.Name = "btnRemoveManager"
		Me.btnRemoveManager.Size = New System.Drawing.Size(99, 30)
		Me.btnRemoveManager.TabIndex = 4
		Me.btnRemoveManager.Text = "Remove"
		Me.btnRemoveManager.UseVisualStyleBackColor = True
		'
		'btnAddManager
		'
		Me.btnAddManager.FlatAppearance.BorderColor = System.Drawing.Color.SteelBlue
		Me.btnAddManager.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
		Me.btnAddManager.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SteelBlue
		Me.btnAddManager.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.btnAddManager.Location = New System.Drawing.Point(6, 213)
		Me.btnAddManager.Name = "btnAddManager"
		Me.btnAddManager.Size = New System.Drawing.Size(57, 30)
		Me.btnAddManager.TabIndex = 3
		Me.btnAddManager.Text = "Add"
		Me.btnAddManager.UseVisualStyleBackColor = True
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.btnRemoveEmp)
		Me.GroupBox2.Controls.Add(Me.btnTransferEmp)
		Me.GroupBox2.Controls.Add(Me.cmbTransferManager)
		Me.GroupBox2.Controls.Add(Me.cmbToManager)
		Me.GroupBox2.Controls.Add(Me.Label6)
		Me.GroupBox2.Controls.Add(Me.Label3)
		Me.GroupBox2.Controls.Add(Me.btnAddEmployee)
		Me.GroupBox2.Controls.Add(Me.cmbRemoveEmployee)
		Me.GroupBox2.Controls.Add(Me.txtAddEmployee)
		Me.GroupBox2.Controls.Add(Me.cmbCurrentEmployee)
		Me.GroupBox2.Controls.Add(Me.Label4)
		Me.GroupBox2.Controls.Add(Me.Label8)
		Me.GroupBox2.Controls.Add(Me.Label2)
		Me.GroupBox2.Location = New System.Drawing.Point(193, 256)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(368, 268)
		Me.GroupBox2.TabIndex = 16
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Admin: Employee"
		'
		'btnRemoveEmp
		'
		Me.btnRemoveEmp.FlatAppearance.BorderColor = System.Drawing.Color.SteelBlue
		Me.btnRemoveEmp.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
		Me.btnRemoveEmp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SteelBlue
		Me.btnRemoveEmp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.btnRemoveEmp.Location = New System.Drawing.Point(240, 213)
		Me.btnRemoveEmp.Name = "btnRemoveEmp"
		Me.btnRemoveEmp.Size = New System.Drawing.Size(117, 30)
		Me.btnRemoveEmp.TabIndex = 7
		Me.btnRemoveEmp.Text = "Remove Employee"
		Me.btnRemoveEmp.UseVisualStyleBackColor = True
		'
		'btnTransferEmp
		'
		Me.btnTransferEmp.FlatAppearance.BorderColor = System.Drawing.Color.SteelBlue
		Me.btnTransferEmp.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
		Me.btnTransferEmp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SteelBlue
		Me.btnTransferEmp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.btnTransferEmp.Location = New System.Drawing.Point(120, 213)
		Me.btnTransferEmp.Name = "btnTransferEmp"
		Me.btnTransferEmp.Size = New System.Drawing.Size(114, 30)
		Me.btnTransferEmp.TabIndex = 6
		Me.btnTransferEmp.Text = "Transfer Employee"
		Me.btnTransferEmp.UseVisualStyleBackColor = True
		'
		'cmbTransferManager
		'
		Me.cmbTransferManager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbTransferManager.FormattingEnabled = True
		Me.cmbTransferManager.Location = New System.Drawing.Point(198, 46)
		Me.cmbTransferManager.Name = "cmbTransferManager"
		Me.cmbTransferManager.Size = New System.Drawing.Size(162, 23)
		Me.cmbTransferManager.Sorted = True
		Me.cmbTransferManager.TabIndex = 3
		'
		'cmbToManager
		'
		Me.cmbToManager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbToManager.FormattingEnabled = True
		Me.cmbToManager.Location = New System.Drawing.Point(198, 108)
		Me.cmbToManager.Name = "cmbToManager"
		Me.cmbToManager.Size = New System.Drawing.Size(162, 23)
		Me.cmbToManager.Sorted = True
		Me.cmbToManager.TabIndex = 4
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Location = New System.Drawing.Point(195, 90)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(84, 15)
		Me.Label3.TabIndex = 17
		Me.Label3.Text = "New Manager:"
		'
		'btnAddEmployee
		'
		Me.btnAddEmployee.FlatAppearance.BorderColor = System.Drawing.Color.SteelBlue
		Me.btnAddEmployee.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
		Me.btnAddEmployee.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SteelBlue
		Me.btnAddEmployee.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.btnAddEmployee.Location = New System.Drawing.Point(9, 213)
		Me.btnAddEmployee.Name = "btnAddEmployee"
		Me.btnAddEmployee.Size = New System.Drawing.Size(105, 30)
		Me.btnAddEmployee.TabIndex = 5
		Me.btnAddEmployee.Text = "Add Employee"
		Me.btnAddEmployee.UseVisualStyleBackColor = True
		'
		'cmbRemoveEmployee
		'
		Me.cmbRemoveEmployee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbRemoveEmployee.FormattingEnabled = True
		Me.cmbRemoveEmployee.Location = New System.Drawing.Point(6, 170)
		Me.cmbRemoveEmployee.Name = "cmbRemoveEmployee"
		Me.cmbRemoveEmployee.Size = New System.Drawing.Size(162, 23)
		Me.cmbRemoveEmployee.Sorted = True
		Me.cmbRemoveEmployee.TabIndex = 2
		'
		'txtAddEmployee
		'
		Me.txtAddEmployee.Location = New System.Drawing.Point(6, 108)
		Me.txtAddEmployee.Name = "txtAddEmployee"
		Me.txtAddEmployee.Size = New System.Drawing.Size(162, 23)
		Me.txtAddEmployee.TabIndex = 1
		'
		'cmbCurrentEmployee
		'
		Me.cmbCurrentEmployee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbCurrentEmployee.FormattingEnabled = True
		Me.cmbCurrentEmployee.Location = New System.Drawing.Point(6, 46)
		Me.cmbCurrentEmployee.Name = "cmbCurrentEmployee"
		Me.cmbCurrentEmployee.Size = New System.Drawing.Size(162, 23)
		Me.cmbCurrentEmployee.Sorted = True
		Me.cmbCurrentEmployee.TabIndex = 0
		'
		'TextBox1
		'
		Me.TextBox1.Location = New System.Drawing.Point(12, 12)
		Me.TextBox1.Multiline = True
		Me.TextBox1.Name = "TextBox1"
		Me.TextBox1.ReadOnly = True
		Me.TextBox1.Size = New System.Drawing.Size(549, 238)
		Me.TextBox1.TabIndex = 17
		Me.TextBox1.Text = resources.GetString("TextBox1.Text")
		'
		'frmUpdate
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.Color.White
		Me.ClientSize = New System.Drawing.Size(576, 538)
		Me.Controls.Add(Me.TextBox1)
		Me.Controls.Add(Me.GroupBox2)
		Me.Controls.Add(Me.GroupBox1)
		Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.Name = "frmUpdate"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "Admin"
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		Me.GroupBox2.ResumeLayout(False)
		Me.GroupBox2.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents Label1 As Label
	Friend WithEvents Label2 As Label
	Friend WithEvents Label4 As Label
	Friend WithEvents Label5 As Label
	Friend WithEvents Label6 As Label
	Friend WithEvents Label7 As Label
	Friend WithEvents Label8 As Label
	Friend WithEvents cmbCurrentManager As ComboBox
	Friend WithEvents cmbRemoveManager As ComboBox
	Friend WithEvents txtAddManager As TextBox
	Friend WithEvents GroupBox1 As GroupBox
	Friend WithEvents btnAddManager As Button
	Friend WithEvents GroupBox2 As GroupBox
	Friend WithEvents cmbTransferManager As ComboBox
	Friend WithEvents cmbToManager As ComboBox
	Friend WithEvents Label3 As Label
	Friend WithEvents btnAddEmployee As Button
	Friend WithEvents cmbRemoveEmployee As ComboBox
	Friend WithEvents txtAddEmployee As TextBox
	Friend WithEvents cmbCurrentEmployee As ComboBox
	Friend WithEvents TextBox1 As TextBox
	Friend WithEvents btnRemoveManager As Button
	Friend WithEvents btnRemoveEmp As Button
	Friend WithEvents btnTransferEmp As Button
End Class
