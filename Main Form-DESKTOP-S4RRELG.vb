'Allstate Service Delivery Coaching Tool
'Developed by Chris DePaul using VB.NET
'Version 1.0 created on 8/12/2021

Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Populate the combo box for employees (Team DePaul) upon the form loading.
        cmbEmployee.Items.Add("Alcivar, Denni")
        cmbEmployee.Items.Add("Blades, Alex")
        cmbEmployee.Items.Add("Cook, Latoya")
        cmbEmployee.Items.Add("Cooper, DeBaron")
        cmbEmployee.Items.Add("Filonov, Daniel")
        cmbEmployee.Items.Add("Fox, Carla")
        cmbEmployee.Items.Add("Garcia, Alicia")
        cmbEmployee.Items.Add("Heesch, David")
        cmbEmployee.Items.Add("Hemric, John")
        cmbEmployee.Items.Add("Jones, Christina")
        cmbEmployee.Items.Add("Keaton, DeeDee")
        cmbEmployee.Items.Add("Kletter, Britney")
        cmbEmployee.Items.Add("Koteles, Liz")
        cmbEmployee.Items.Add("Margeson, Karen")
        cmbEmployee.Items.Add("McKamie, Kirsten")
        cmbEmployee.Items.Add("Power, Melissa")
        cmbEmployee.Items.Add("Ramirez, Josslin")
        cmbEmployee.Items.Add("Shelley, Jenna")
        cmbEmployee.Items.Add("Steve, Theresa")

        'Populate the combo box for the month selection with 12 months.
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
        'Open TalentConnection main page

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
        Dim strAllstate As String = "https://home.allstate.com"
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

    Private Sub mnuHelp_Click(sender As Object, e As EventArgs) Handles mnuHelp.Click
        'Show the help form upon clicking the Help button.
        frmHelp.Show()

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear the textbox input upon clicking the Clear button and also clear the Employee Selection and Month Selection combo boxes.

        'Clear combo boxes.
        cmbEmployee.SelectedIndex = -1
        cmbMonth.SelectedIndex = -1

        'Clear text boxes.
        txtDuration.Clear()
        txtCSAT.Clear()
        txtAway.Clear()
        txtQual.Clear()
        txtDev.Clear()

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Determine the month and employee from the selection combo boxes and save the information.

        'Delcare variables as doubles for the metric input.
        Dim dblHCOPH As Double = txtDuration.Text
        Dim dblCSAT As Double = txtCSAT.Text
        Dim dblAway As Double = txtAway.Text
        Dim dblQual As Double = txtQual.Text
        Dim dblDev As Double = txtDev.Text
        Dim strEmployee As String = cmbEmployee.Text
        Dim strMonth As String = cmbMonth.Text

        'Parse the string contents of the textboxes into a double.
        Double.TryParse(txtDuration.Text, dblHCOPH)
        Double.TryParse(txtCSAT.Text, dblCSAT)
        Double.TryParse(txtAway.Text, dblAway)
        Double.TryParse(txtQual.Text, dblQual)
        Double.TryParse(txtDev.Text, dblDev)

    End Sub

End Class
