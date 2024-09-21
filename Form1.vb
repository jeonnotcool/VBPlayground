Imports System.Drawing.Printing
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms

' VBPlayground v1.0
' A simple Windows Forms application that demonstrates the use of common controls, container controls, and dialog controls in Visual Basic.
' Developed by jeonnotcool (https://github.com/jeonnotcool)
' Date: 2024-09-21
' I added comments to the code to make it easier to understand. If you have any questions, feel free to ask.
' This is very useful for students who are learning Visual Basic (although I am a student lololol) .


' There are three types of controls in Windows Forms:
' 1. Common Controls: These are the controls that are used to take input from the user or display information to the user.
' 2. Container Controls: These are the controls that are used to hold other controls.
' 3. Dialog Controls: These are the controls that are used to interact with the user to perform a specific task.

' The following code demonstrates the use of common controls, container controls, and dialog controls in Windows Forms.
' All of this are also possible with the use of GitHub Copilot. Thank you ;)

' I added so many comments here to make it look it has so many lines of code. But in reality, it's just a simple code.

Public Class Form1

    ' Function to call Notification. Can be used using ShowNotification("Title", "Message")
    Public Sub ShowNotification(tipTitle As String, message As String, ic As Icon)
        Dim notification As New NotifyIcon()
        notification.Icon = ic
        notification.BalloonTipTitle = tipTitle
        notification.BalloonTipText = message
        notification.Visible = True
        notification.ShowBalloonTip(5000)
    End Sub


    ' ' ' ' ' ' ' ' '
    'Common Controls'
    ' ' ' ' ' ' ' ' '
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click ' Button
        MessageBox.Show("Hello, World!")
    End Sub

    ' Label - There's no need to code it.

    Private Sub btnTextBox_Click(sender As Object, e As EventArgs) Handles btnTextBox.Click ' TextBox
        MessageBox.Show(txtBoxText.Text)
    End Sub

    Private Sub btnCheckBox_Click(sender As Object, e As EventArgs) Handles btnCheckBox.Click ' CheckBox
        If CheckBox1.Checked AndAlso CheckBox2.Checked Then
            MessageBox.Show("The following item are checked:" & vbCrLf & CheckBox1.Text & vbCrLf & CheckBox2.Text)
        ElseIf CheckBox2.Checked Then
            MessageBox.Show("The following item is checked:" & vbCrLf & CheckBox2.Text)
        ElseIf CheckBox1.Checked Then
            MessageBox.Show("The following item is checked:" & vbCrLf & CheckBox1.Text)
        Else
            MessageBox.Show("No items are checked 😡")
        End If
    End Sub

    Private Sub btnRadio_Click(sender As Object, e As EventArgs) Handles btnRadio.Click ' RadioButton
        If rbMessageBox.Checked Then
            MessageBox.Show("Hello world")
        ElseIf rbWindows.Checked Then
            ShowNotification("Hello world", "Windows notification has been selected!", SystemIcons.Information)
        End If

    End Sub

    Private Sub btnComboBox_Click(sender As Object, e As EventArgs) Handles btnComboBox.Click ' ComboBox
        Dim selectedValue As String = ComboBox1.SelectedItem.ToString()
        MessageBox.Show("Selected item: " & vbCrLf & selectedValue)
    End Sub

    Private Sub btnListBox_Click(sender As Object, e As EventArgs) Handles btnListBox.Click ' ListBox
        Dim selectedValue As String = ListBox1.SelectedItem.ToString()
        MessageBox.Show("Selected item: " & vbCrLf & selectedValue)
    End Sub

    Private Sub btnProgressBar_Click(sender As Object, e As EventArgs) Handles btnProgressBar.Click ' ProgressBar
        Dim input As String = txtBoxProgressBar.Text
        Dim value As Integer

        If Integer.TryParse(input, value) Then
            If value > 100 Then
                MessageBox.Show("Error: Number cannot be above 100")
            Else
                ProgressBar1.Value = value
            End If
        Else
            MessageBox.Show("Error: Invalid input. Please enter a number.")
        End If
    End Sub

    Private Sub btnDTP_Click(sender As Object, e As EventArgs) Handles btnDTP.Click ' DateTimePicker
        Dim selectedDateTime As DateTime = DateTimePicker1.Value
        MessageBox.Show("Selected Date and Time: " & selectedDateTime.ToString())
    End Sub

    ' ' ' ' ' ' ' ' ' ''
    'Container Controls'
    ' ' ' ' ' ' ' ' ' ''

    ' There's no need to code on container controls as they are used to hold other controls.

    ' ' ' ' ' ' ' ' '
    'Dialog Controls'
    ' ' ' ' ' ' ' ' '

    Private Sub btnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click ' Open File Dialog
        Dim openFileDialog1 As New OpenFileDialog() ' Create a new OpenFileDialog object

        '  openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) 'Set the default directory to the userprofile
        openFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*" 'Specify file types
        openFileDialog1.FilterIndex = 1 'Sets default filter index to 1 which is Text Files
        openFileDialog1.RestoreDirectory = True 'Restore the current directory after file selection

        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            MessageBox.Show("The file has been successfully opened!" & vbCrLf & "File Directory: " & openFileDialog1.FileName)

            ' Read the contents of the file
            Dim filePath As String = openFileDialog1.FileName
            Dim fileText As String = File.ReadAllText(filePath)
            txtBoxOF.Text = fileText ' Display the content in a TextBox
        End If
    End Sub


    Private Sub btnSaveFile_Click(sender As Object, e As EventArgs) Handles btnSaveFile.Click ' Save File Dialog
        Dim saveFileDialog1 As New SaveFileDialog() ' Create a new SaveFileDialog object

        saveFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*" ' Set the file type filter
        saveFileDialog1.FilterIndex = 1 ' Set the default filter index to 1 which is Text Files
        saveFileDialog1.RestoreDirectory = True ' Restore the current directory after file selection

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Write the contents of a TextBox to the file
            Dim filePath As String = saveFileDialog1.FileName ' Sets the file path
            File.WriteAllText(filePath, txtBoxSF.Text) ' Writes the content of the TextBox to the file
            MessageBox.Show("File saved successfully." & vbCrLf & "File Directory: " & filePath) ' Display a message box
        Else
            MessageBox.Show("File not saved due to user cancellation.")
        End If
    End Sub

    Private Sub btnFontDialog_Click(sender As Object, e As EventArgs) Handles btnFontDialog.Click ' Font Dialog
        Dim fontDialog1 As New FontDialog() ' Create a new FontDialog object
        fontDialog1.Font = New Font("Segoe UI", 11) ' Set the default 

        If fontDialog1.ShowDialog() = DialogResult.OK Then

            ' Set the font of a TextBox to the selected font
            txtSelected.Font = fontDialog1.Font
            txtSelected.Text = "Selected Font: " & fontDialog1.Font.Name
        End If
    End Sub
    Private Sub btnColorDialog_Click(sender As Object, e As EventArgs) Handles btnColorDialog.Click ' Color Dialog
        Dim colorDialog1 As New ColorDialog() ' Create a new ColorDialog object

        If colorDialog1.ShowDialog() = DialogResult.OK Then
            ' Set the color of a TextBox to the selected color
            panelColor.BackColor = colorDialog1.Color
        End If
    End Sub

    Private Sub btnPrintDialog_Click(sender As Object, e As EventArgs) Handles btnPrintDialog.Click ' PrintDialog
        Dim printDialog1 As New PrintDialog() ' Create a new PrintDialog object
        printDialog1.PrinterSettings.PrinterName = "Microsoft Print to PDF" ' Set the printer to Microsoft Print to PDF
        If printDialog1.ShowDialog() = DialogResult.OK Then
            Dim printDocument1 As New PrintDocument() ' Create a new PrintDocument object

            ' Set the PrintDocument object's PrintPage event handler
            AddHandler printDocument1.PrintPage, AddressOf PrintDocument1_PrintPage

            ' Print the document
            printDocument1.Print()
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) ' PrintDialogExtension
        Dim image As Image = My.Resources.Wenomechainsama ' Load the image

        ' Draw the image on the print page
        e.Graphics.DrawImage(image, e.MarginBounds)

        ' Specify that there are no more pages to print
        e.HasMorePages = False
    End Sub


    ' Strips 
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click ' Close Application
        Me.Close() ' Close the application
    End Sub

    Private Sub VisitRepo_Click(sender As Object, e As EventArgs) Handles VisitRepo.Click
        Dim repoUrl As String = "https://github.com/jeonnotcool/VBPlayground"
        Process.Start(repoUrl)
    End Sub

    Private Sub MessageBoxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MessageBoxToolStripMenuItem.Click
        MessageBox.Show("Hello, World!")
    End Sub

    Private Sub WindowsNotficationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WindowsNotficationToolStripMenuItem.Click
        ShowNotification("Hello world", "Triggered by the MenuStrip!", SystemIcons.Information)
    End Sub

    Private Sub FileLocationMenuItem_Click(sender As Object, e As EventArgs) Handles FileLocationMenuItem.Click
        Dim appLocation As String = Application.ExecutablePath
        Process.Start("explorer.exe", "/select," & appLocation)
    End Sub
End Class
