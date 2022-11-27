<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdmin
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
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAdmin))
		Me.btnLogin = New System.Windows.Forms.Button()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.txtUsername = New System.Windows.Forms.TextBox()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.txtPassword = New System.Windows.Forms.TextBox()
		Me.SuspendLayout()
		'
		'btnLogin
		'
		Me.btnLogin.FlatAppearance.BorderColor = System.Drawing.Color.SteelBlue
		Me.btnLogin.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
		Me.btnLogin.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SteelBlue
		Me.btnLogin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.btnLogin.Location = New System.Drawing.Point(116, 67)
		Me.btnLogin.Name = "btnLogin"
		Me.btnLogin.Size = New System.Drawing.Size(85, 32)
		Me.btnLogin.TabIndex = 0
		Me.btnLogin.Text = "&Login"
		Me.btnLogin.UseVisualStyleBackColor = True
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(32, 9)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(63, 15)
		Me.Label1.TabIndex = 1
		Me.Label1.Text = "Username:"
		'
		'txtUsername
		'
		Me.txtUsername.Location = New System.Drawing.Point(101, 6)
		Me.txtUsername.Name = "txtUsername"
		Me.txtUsername.Size = New System.Drawing.Size(115, 23)
		Me.txtUsername.TabIndex = 2
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(35, 41)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(60, 15)
		Me.Label2.TabIndex = 3
		Me.Label2.Text = "Password:"
		'
		'txtPassword
		'
		Me.txtPassword.Location = New System.Drawing.Point(101, 38)
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.Size = New System.Drawing.Size(115, 23)
		Me.txtPassword.TabIndex = 4
		'
		'frmAdmin
		'
		Me.AcceptButton = Me.btnLogin
		Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.Color.White
		Me.ClientSize = New System.Drawing.Size(270, 117)
		Me.Controls.Add(Me.txtPassword)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.txtUsername)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.btnLogin)
		Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.Name = "frmAdmin"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "Admin Login"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents btnLogin As Button
	Friend WithEvents Label1 As Label
	Friend WithEvents txtUsername As TextBox
	Friend WithEvents Label2 As Label
	Friend WithEvents txtPassword As TextBox
End Class
