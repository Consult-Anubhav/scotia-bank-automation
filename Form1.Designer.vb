<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnK2 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnRootPath = New System.Windows.Forms.Button()
        Me.FakeRootPathLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnK2
        '
        Me.btnK2.Location = New System.Drawing.Point(42, 126)
        Me.btnK2.Name = "btnK2"
        Me.btnK2.Size = New System.Drawing.Size(75, 23)
        Me.btnK2.TabIndex = 0
        Me.btnK2.Text = "K2"
        Me.btnK2.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnRootPath
        '
        Me.btnRootPath.Location = New System.Drawing.Point(42, 32)
        Me.btnRootPath.Name = "btnRootPath"
        Me.btnRootPath.Size = New System.Drawing.Size(102, 23)
        Me.btnRootPath.TabIndex = 1
        Me.btnRootPath.Text = "Select RootPath"
        Me.btnRootPath.UseVisualStyleBackColor = True
        '
        'FakeRootPathLabel
        '
        Me.FakeRootPathLabel.AutoSize = True
        Me.FakeRootPathLabel.Location = New System.Drawing.Point(150, 41)
        Me.FakeRootPathLabel.Name = "FakeRootPathLabel"
        Me.FakeRootPathLabel.Size = New System.Drawing.Size(0, 13)
        Me.FakeRootPathLabel.TabIndex = 2
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.FakeRootPathLabel)
        Me.Controls.Add(Me.btnRootPath)
        Me.Controls.Add(Me.btnK2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnK2 As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnRootPath As Button
    Friend WithEvents FakeRootPathLabel As Label
End Class
