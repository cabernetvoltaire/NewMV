<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MovieSwapTest
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MovieSwapTest))
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Test1 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.Test2 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.Test1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Test2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 24
        Me.ListBox1.Location = New System.Drawing.Point(2193, 55)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(388, 844)
        Me.ListBox1.TabIndex = 0
        '
        'Test1
        '
        Me.Test1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Test1.Enabled = True
        Me.Test1.Location = New System.Drawing.Point(0, 0)
        Me.Test1.Margin = New System.Windows.Forms.Padding(4)
        Me.Test1.Name = "Test1"
        Me.Test1.OcxState = CType(resources.GetObject("Test1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Test1.Size = New System.Drawing.Size(2164, 1317)
        Me.Test1.TabIndex = 6
        Me.Test1.TabStop = False
        Me.Test1.UseWaitCursor = True
        '
        'Test2
        '
        Me.Test2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Test2.Enabled = True
        Me.Test2.Location = New System.Drawing.Point(0, 0)
        Me.Test2.Margin = New System.Windows.Forms.Padding(4)
        Me.Test2.Name = "Test2"
        Me.Test2.OcxState = CType(resources.GetObject("Test2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Test2.Size = New System.Drawing.Size(2164, 1317)
        Me.Test2.TabIndex = 5
        Me.Test2.TabStop = False
        Me.Test2.UseWaitCursor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Test2)
        Me.Panel1.Controls.Add(Me.Test1)
        Me.Panel1.Location = New System.Drawing.Point(23, 37)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(2164, 1317)
        Me.Panel1.TabIndex = 7
        '
        'MovieSwapTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2733, 1312)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "MovieSwapTest"
        Me.Text = "MovieSwapTest"
        CType(Me.Test1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Test2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Test1 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents Test2 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents Panel1 As Panel
End Class
