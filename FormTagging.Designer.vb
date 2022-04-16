<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormTagging
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
        Me.components = New System.ComponentModel.Container()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.MVDataSet = New MasaSam.Forms.Sample.MVDataSet()
        Me.MarksNoCryptBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MarksNoCryptTableAdapter = New MasaSam.Forms.Sample.MVDataSetTableAdapters.MarksNoCryptTableAdapter()
        CType(Me.MVDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MarksNoCryptBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.DataSource = Me.MarksNoCryptBindingSource
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 20
        Me.ListBox1.Location = New System.Drawing.Point(12, 143)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(776, 284)
        Me.ListBox1.TabIndex = 0
        '
        'MVDataSet
        '
        Me.MVDataSet.DataSetName = "MVDataSet"
        Me.MVDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'MarksNoCryptBindingSource
        '
        Me.MarksNoCryptBindingSource.DataMember = "MarksNoCrypt"
        Me.MarksNoCryptBindingSource.DataSource = Me.MVDataSet
        '
        'MarksNoCryptTableAdapter
        '
        Me.MarksNoCryptTableAdapter.ClearBeforeFill = True
        '
        'FormTagging
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "FormTagging"
        Me.Text = "FormTagging"
        CType(Me.MVDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MarksNoCryptBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents MVDataSet As MVDataSet
    Friend WithEvents MarksNoCryptBindingSource As BindingSource
    Friend WithEvents MarksNoCryptTableAdapter As MVDataSetTableAdapters.MarksNoCryptTableAdapter
End Class
