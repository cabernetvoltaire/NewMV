<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FullScreen
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FullScreen))
        Me.FSWMP = New AxWMPLib.AxWindowsMediaPlayer()
        Me.FSBlanker = New System.Windows.Forms.PictureBox()
        Me.FSPB1 = New System.Windows.Forms.PictureBox()
        Me.DirectoryEntry1 = New System.DirectoryServices.DirectoryEntry()
        Me.FSWMP2 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.FSWMP3 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.FSPB2 = New System.Windows.Forms.PictureBox()
        Me.FSPB3 = New System.Windows.Forms.PictureBox()
        CType(Me.FSWMP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSBlanker, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSPB1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSWMP2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSWMP3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSPB2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSPB3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FSWMP
        '
        Me.FSWMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FSWMP.Enabled = True
        Me.FSWMP.Location = New System.Drawing.Point(0, 0)
        Me.FSWMP.Margin = New System.Windows.Forms.Padding(2)
        Me.FSWMP.Name = "FSWMP"
        Me.FSWMP.OcxState = CType(resources.GetObject("FSWMP.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FSWMP.Size = New System.Drawing.Size(1571, 900)
        Me.FSWMP.TabIndex = 0
        Me.FSWMP.TabStop = False
        Me.FSWMP.UseWaitCursor = True
        '
        'FSBlanker
        '
        Me.FSBlanker.BackColor = System.Drawing.Color.Maroon
        Me.FSBlanker.Location = New System.Drawing.Point(322, 258)
        Me.FSBlanker.Margin = New System.Windows.Forms.Padding(2)
        Me.FSBlanker.Name = "FSBlanker"
        Me.FSBlanker.Size = New System.Drawing.Size(976, 550)
        Me.FSBlanker.TabIndex = 19
        Me.FSBlanker.TabStop = False
        Me.FSBlanker.Visible = False
        '
        'FSPB1
        '
        Me.FSPB1.BackColor = System.Drawing.Color.Black
        Me.FSPB1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.FSPB1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FSPB1.Location = New System.Drawing.Point(0, 0)
        Me.FSPB1.Margin = New System.Windows.Forms.Padding(0)
        Me.FSPB1.Name = "FSPB1"
        Me.FSPB1.Size = New System.Drawing.Size(1571, 900)
        Me.FSPB1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.FSPB1.TabIndex = 18
        Me.FSPB1.TabStop = False
        Me.FSPB1.Visible = False
        '
        'FSWMP2
        '
        Me.FSWMP2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FSWMP2.Enabled = True
        Me.FSWMP2.Location = New System.Drawing.Point(0, 0)
        Me.FSWMP2.Margin = New System.Windows.Forms.Padding(2)
        Me.FSWMP2.Name = "FSWMP2"
        Me.FSWMP2.OcxState = CType(resources.GetObject("FSWMP2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FSWMP2.Size = New System.Drawing.Size(1571, 900)
        Me.FSWMP2.TabIndex = 20
        Me.FSWMP2.TabStop = False
        Me.FSWMP2.UseWaitCursor = True
        '
        'FSWMP3
        '
        Me.FSWMP3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FSWMP3.Enabled = True
        Me.FSWMP3.Location = New System.Drawing.Point(0, 0)
        Me.FSWMP3.Margin = New System.Windows.Forms.Padding(2)
        Me.FSWMP3.Name = "FSWMP3"
        Me.FSWMP3.OcxState = CType(resources.GetObject("FSWMP3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FSWMP3.Size = New System.Drawing.Size(1571, 900)
        Me.FSWMP3.TabIndex = 21
        Me.FSWMP3.TabStop = False
        Me.FSWMP3.UseWaitCursor = True
        '
        'FSPB2
        '
        Me.FSPB2.BackColor = System.Drawing.Color.Black
        Me.FSPB2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.FSPB2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FSPB2.Location = New System.Drawing.Point(0, 0)
        Me.FSPB2.Margin = New System.Windows.Forms.Padding(0)
        Me.FSPB2.Name = "FSPB2"
        Me.FSPB2.Size = New System.Drawing.Size(1571, 900)
        Me.FSPB2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.FSPB2.TabIndex = 22
        Me.FSPB2.TabStop = False
        Me.FSPB2.Visible = False
        '
        'FSPB3
        '
        Me.FSPB3.BackColor = System.Drawing.Color.Black
        Me.FSPB3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.FSPB3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FSPB3.Location = New System.Drawing.Point(0, 0)
        Me.FSPB3.Margin = New System.Windows.Forms.Padding(0)
        Me.FSPB3.Name = "FSPB3"
        Me.FSPB3.Size = New System.Drawing.Size(1571, 900)
        Me.FSPB3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.FSPB3.TabIndex = 23
        Me.FSPB3.TabStop = False
        Me.FSPB3.Visible = False
        '
        'FullScreen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(1571, 900)
        Me.Controls.Add(Me.FSPB3)
        Me.Controls.Add(Me.FSPB2)
        Me.Controls.Add(Me.FSWMP3)
        Me.Controls.Add(Me.FSWMP2)
        Me.Controls.Add(Me.FSBlanker)
        Me.Controls.Add(Me.FSPB1)
        Me.Controls.Add(Me.FSWMP)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "FullScreen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.TopMost = True
        CType(Me.FSWMP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSBlanker, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSPB1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSWMP2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSWMP3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSPB2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSPB3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FSWMP As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents FSPB1 As PictureBox
    Friend WithEvents FSBlanker As PictureBox
    Friend WithEvents DirectoryEntry1 As DirectoryServices.DirectoryEntry
    Friend WithEvents FSWMP2 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents FSWMP3 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents FSPB2 As PictureBox
    Friend WithEvents FSPB3 As PictureBox
End Class
