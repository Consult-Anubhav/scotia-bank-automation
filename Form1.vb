Imports System.Net.Mime.MediaTypeNames
Imports System.Threading
Imports Microsoft.Office.Interop.Word

Public Class Form1
    Private Sub btnK2_Click(sender As Object, e As EventArgs)
        TestK2()
    End Sub

    Private Sub btnRootPath_Click(sender As Object, e As EventArgs) Handles btnRootPath.Click
        ' Show a folder browser dialog to select the folder location
        Dim folderBrowserDialog As New FolderBrowserDialog()
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected folder path
            Dim selectedFolderPath As String = folderBrowserDialog.SelectedPath

            ' Assign the selected folder path to fakeRootPath in Faker
            Faker.SetFakeRootPath(selectedFolderPath)

            ' update any UI elements to reflect the selected folder path
            ' FakeRootPathLabel.Text = selectedFolderPath '  FakeRootPathLabel is a label to display the selected folder path
            Dim item = New ToolStripStatusLabel("Root Path: " & selectedFolderPath)


            StatusStrip1.Items.Clear()
            StatusStrip1.Items.Add(item)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = -1 Then

            btnStart.Enabled = False

        Else
            btnStart.Enabled = True
        End If


    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        Select Case ComboBox1.SelectedItem.ToString()
            Case "K2"
                TestK2()
            Case "LATAM"
                DownloadAndProcessForexData()
            Case "OPICS"
                FXCalc()
            Case "SCOTS"
                TestSCOTS()
            Case "Murex"
                TestMurex()

            Case Else
                Console.WriteLine("Invalid grade")
        End Select

    End Sub



    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'FakeRootPathLabel.Text = "Path:" & Faker.GetFakeRootPath() '  FakeRootPathLabel is a label to display the selected folder path
        Dim item = New ToolStripStatusLabel("Root Path: " & Faker.GetFakeRootPath())


        StatusStrip1.Items.Add(item)
        If ComboBox1.SelectedIndex = -1 Then

            btnStart.Enabled = False

        Else
            btnStart.Enabled = True
        End If


    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub


End Class
