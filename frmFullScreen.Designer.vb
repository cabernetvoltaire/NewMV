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
        Me.fullScreenPicBox = New System.Windows.Forms.PictureBox()
        Me.DirectoryEntry1 = New System.DirectoryServices.DirectoryEntry()
        Me.FSWMP2 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.FSWMP3 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        CType(Me.FSWMP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSBlanker, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fullScreenPicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSWMP2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSWMP3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FSWMP
        '
        Me.FSWMP.Enabled = True
        Me.FSWMP.Location = New System.Drawing.Point(0, 0)
        Me.FSWMP.Name = "FSWMP"
        Me.FSWMP.OcxState = CType(resources.GetObject("FSWMP.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FSWMP.Size = New System.Drawing.Size(1123, 720)
        Me.FSWMP.TabIndex = 0
        Me.FSWMP.TabStop = False
        Me.FSWMP.UseWaitCursor = True
        '
        'FSBlanker
        '
        Me.FSBlanker.BackColor = System.Drawing.Color.Maroon
        Me.FSBlanker.Location = New System.Drawing.Point(394, 310)
        Me.FSBlanker.Name = "FSBlanker"
        Me.FSBlanker.Size = New System.Drawing.Size(1193, 660)
        Me.FSBlanker.TabIndex = 19
        Me.FSBlanker.TabStop = False
        Me.FSBlanker.Visible = False
        '
        'fullScreenPicBox
        '
        Me.fullScreenPicBox.BackColor = System.Drawing.Color.Black
        Me.fullScreenPicBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.fullScreenPicBox.Location = New System.Drawing.Point(608, 637)
        Me.fullScreenPicBox.Margin = New System.Windows.Forms.Padding(0)
        Me.fullScreenPicBox.Name = "fullScreenPicBox"
        Me.fullScreenPicBox.Size = New System.Drawing.Size(1920, 1080)
        Me.fullScreenPicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.fullScreenPicBox.TabIndex = 18
        Me.fullScreenPicBox.TabStop = False
        Me.fullScreenPicBox.Visible = False
        '
        'FSWMP2
        '
        Me.FSWMP2.Enabled = True
        Me.FSWMP2.Location = New System.Drawing.Point(399, 180)
        Me.FSWMP2.Name = "FSWMP2"
        Me.FSWMP2.OcxState = CType(resources.GetObject("FSWMP2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FSWMP2.Size = New System.Drawing.Size(1123, 720)
        Me.FSWMP2.TabIndex = 20
        Me.FSWMP2.TabStop = False
        Me.FSWMP2.UseWaitCursor = True
        '
        'FSWMP3
        '
        Me.FSWMP3.Enabled = True
        Me.FSWMP3.Location = New System.Drawing.Point(671, 348)
        Me.FSWMP3.Name = "FSWMP3"
        Me.FSWMP3.OcxState = CType(resources.GetObject("FSWMP3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FSWMP3.Size = New System.Drawing.Size(1123, 720)
        Me.FSWMP3.TabIndex = 21
        Me.FSWMP3.TabStop = False
        Me.FSWMP3.UseWaitCursor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Black
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1920, 1080)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Black
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Location = New System.Drawing.Point(16, 16)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(0)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(1920, 1080)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox2.TabIndex = 23
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Visible = False
        '
        'FullScreen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(1920, 1080)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.FSWMP3)
        Me.Controls.Add(Me.FSWMP2)
        Me.Controls.Add(Me.FSBlanker)
        Me.Controls.Add(Me.fullScreenPicBox)
        Me.Controls.Add(Me.FSWMP)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Name = "FullScreen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.TopMost = True
        CType(Me.FSWMP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSBlanker, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fullScreenPicBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSWMP2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSWMP3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FSWMP As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents fullScreenPicBox As PictureBox
    Friend WithEvents FSBlanker As PictureBox
    Friend WithEvents DirectoryEntry1 As DirectoryServices.DirectoryEntry
    Friend WithEvents FSWMP2 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents FSWMP3 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents PictureBox2 As PictureBox
End Class
