<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TVPopup
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
        Me.FileSystemTree1 = New MasaSam.Forms.Controls.FileSystemTree()
        Me.SuspendLayout()
        '
        'FileSystemTree1
        '
        Me.FileSystemTree1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FileSystemTree1.FileExtensions = "*"
        Me.FileSystemTree1.Location = New System.Drawing.Point(0, 0)
        Me.FileSystemTree1.Margin = New System.Windows.Forms.Padding(6)
        Me.FileSystemTree1.Name = "FileSystemTree1"
        Me.FileSystemTree1.RootDrive = Nothing
        Me.FileSystemTree1.SelectedFolder = Nothing
        Me.FileSystemTree1.Size = New System.Drawing.Size(343, 602)
        Me.FileSystemTree1.TabIndex = 0
        '
        'TVPopup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(343, 602)
        Me.Controls.Add(Me.FileSystemTree1)
        Me.Name = "TVPopup"
        Me.Text = "TreeView"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FileSystemTree1 As Controls.FileSystemTree
End Class
