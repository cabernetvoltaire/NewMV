<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StartpointTester
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
        Me.FileName = New System.Windows.Forms.Label()
        Me.Duration = New System.Windows.Forms.Label()
        Me.State = New System.Windows.Forms.Label()
        Me.Startpoint = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'FileName
        '
        Me.FileName.AutoSize = True
        Me.FileName.Location = New System.Drawing.Point(44, 27)
        Me.FileName.Name = "FileName"
        Me.FileName.Size = New System.Drawing.Size(57, 20)
        Me.FileName.TabIndex = 0
        Me.FileName.Text = "Label1"
        '
        'Duration
        '
        Me.Duration.AutoSize = True
        Me.Duration.Location = New System.Drawing.Point(44, 72)
        Me.Duration.Name = "Duration"
        Me.Duration.Size = New System.Drawing.Size(57, 20)
        Me.Duration.TabIndex = 1
        Me.Duration.Text = "Label1"
        '
        'State
        '
        Me.State.AutoSize = True
        Me.State.Location = New System.Drawing.Point(44, 119)
        Me.State.Name = "State"
        Me.State.Size = New System.Drawing.Size(57, 20)
        Me.State.TabIndex = 2
        Me.State.Text = "Label1"
        '
        'Startpoint
        '
        Me.Startpoint.AutoSize = True
        Me.Startpoint.Location = New System.Drawing.Point(44, 162)
        Me.Startpoint.Name = "Startpoint"
        Me.Startpoint.Size = New System.Drawing.Size(57, 20)
        Me.Startpoint.TabIndex = 3
        Me.Startpoint.Text = "Label1"
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1585, 462)
        Me.Controls.Add(Me.Startpoint)
        Me.Controls.Add(Me.State)
        Me.Controls.Add(Me.Duration)
        Me.Controls.Add(Me.FileName)
        Me.Name = "Form3"
        Me.Text = "Form3"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents FileName As Label
    Friend WithEvents Duration As Label
    Friend WithEvents State As Label
    Friend WithEvents Startpoint As Label
End Class
