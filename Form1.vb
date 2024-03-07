Public Class Form1
    Private Sub btnK2_Click(sender As Object, e As EventArgs) Handles btnK2.Click
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

            ' Optionally, update any UI elements to reflect the selected folder path
            FakeRootPathLabel.Text = selectedFolderPath '  FakeRootPathLabel is a label to display the selected folder path
        End If
    End Sub
End Class
