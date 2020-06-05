<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DuplicatesForm
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
        Me.DuplicatesTLP = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DuplicatesTLP
        '
        Me.DuplicatesTLP.AutoScroll = True
        Me.DuplicatesTLP.AutoSize = True
        Me.DuplicatesTLP.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.DuplicatesTLP.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DuplicatesTLP.ColumnCount = 1
        Me.DuplicatesTLP.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.DuplicatesTLP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DuplicatesTLP.Location = New System.Drawing.Point(3, 92)
        Me.DuplicatesTLP.Name = "DuplicatesTLP"
        Me.DuplicatesTLP.RowCount = 1
        Me.DuplicatesTLP.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.DuplicatesTLP.Size = New System.Drawing.Size(775, 799)
        Me.DuplicatesTLP.TabIndex = 3
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.BackColor = System.Drawing.Color.IndianRed
        Me.TableLayoutPanel2.ColumnCount = 3
        Me.TableLayoutPanel1.SetColumnSpan(Me.TableLayoutPanel2, 2)
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.Controls.Add(Me.ComboBox2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.ComboBox1, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Button1, 2, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 83.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(1513, 83)
        Me.TableLayoutPanel2.TabIndex = 1
        '
        'ComboBox2
        '
        Me.ComboBox2.DisplayMember = "0"
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"0", "50", "100", "200"})
        Me.ComboBox2.Location = New System.Drawing.Point(3, 3)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(94, 32)
        Me.ComboBox2.TabIndex = 2
        '
        'ComboBox1
        '
        Me.ComboBox1.DisplayMember = "10"
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"10", "50", "100", "200", "All"})
        Me.ComboBox1.Location = New System.Drawing.Point(103, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(94, 32)
        Me.ComboBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(203, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(361, 40)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Search for duplicates"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.AutoScroll = True
        Me.TableLayoutPanel1.AutoSize = True
        Me.TableLayoutPanel1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 738.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.DuplicatesTLP, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel4, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 90.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1519, 894)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 1
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.Controls.Add(Me.ListBox2, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.ListBox1, 0, 1)
        Me.TableLayoutPanel4.Controls.Add(Me.Panel1, 0, 2)
        Me.TableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(784, 92)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 4
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 244.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(732, 799)
        Me.TableLayoutPanel4.TabIndex = 4
        '
        'ListBox2
        '
        Me.ListBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.ItemHeight = 24
        Me.ListBox2.Location = New System.Drawing.Point(3, 3)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.ListBox2.Size = New System.Drawing.Size(726, 244)
        Me.ListBox2.TabIndex = 1
        '
        'ListBox1
        '
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 24
        Me.ListBox1.Location = New System.Drawing.Point(3, 253)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.ListBox1.Size = New System.Drawing.Size(726, 244)
        Me.ListBox1.Sorted = True
        Me.ListBox1.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 503)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(726, 238)
        Me.Panel1.TabIndex = 4
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(254, 0)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(245, 60)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Remove &All (Takes time!)"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(3, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(245, 57)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Remove &Unselected"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DuplicatesForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1519, 894)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "DuplicatesForm"
        Me.Text = "DuplicatesForm"
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DuplicatesTLP As TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Button3 As Button
    Friend WithEvents ComboBox2 As ComboBox
End Class
