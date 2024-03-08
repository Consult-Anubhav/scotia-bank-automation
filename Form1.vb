Imports System.Net.Mime.MediaTypeNames
Imports System.Threading
Imports Microsoft.Office.Interop.Word

Public Class Form1



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
            btnStart.BackColor = Color.DodgerBlue
        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = -1 Then

            btnStart.Enabled = False

        Else
            btnStart.Enabled = True
            btnStart.BackColor = Color.DodgerBlue
        End If
    End Sub
End Class
