<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form4
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
        Me.components = New System.ComponentModel.Container()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.MVDataSet = New MasaSam.Forms.Sample.MVDataSet()
        Me.MVDataSetBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MarksNoCryptBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MarksNoCryptTableAdapter = New MasaSam.Forms.Sample.MVDataSetTableAdapters.MarksNoCryptTableAdapter()
        Me.IDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FFileNameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FPathDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FSizeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FCreateDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FTagsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FMarkersDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FThumbnailDataGridViewImageColumn = New System.Windows.Forms.DataGridViewImageColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MVDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MVDataSetBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MarksNoCryptBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IDDataGridViewTextBoxColumn, Me.FFileNameDataGridViewTextBoxColumn, Me.FPathDataGridViewTextBoxColumn, Me.FSizeDataGridViewTextBoxColumn, Me.FCreateDateDataGridViewTextBoxColumn, Me.FTagsDataGridViewTextBoxColumn, Me.FMarkersDataGridViewTextBoxColumn, Me.FThumbnailDataGridViewImageColumn})
        Me.DataGridView1.DataSource = Me.MarksNoCryptBindingSource
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 31
        Me.DataGridView1.Size = New System.Drawing.Size(978, 540)
        Me.DataGridView1.TabIndex = 0
        '
        'MVDataSet
        '
        Me.MVDataSet.DataSetName = "MVDataSet"
        Me.MVDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'MVDataSetBindingSource
        '
        Me.MVDataSetBindingSource.DataSource = Me.MVDataSet
        Me.MVDataSetBindingSource.Position = 0
        '
        'MarksNoCryptBindingSource
        '
        Me.MarksNoCryptBindingSource.DataMember = "MarksNoCrypt"
        Me.MarksNoCryptBindingSource.DataSource = Me.MVDataSetBindingSource
        '
        'MarksNoCryptTableAdapter
        '
        Me.MarksNoCryptTableAdapter.ClearBeforeFill = True
        '
        'IDDataGridViewTextBoxColumn
        '
        Me.IDDataGridViewTextBoxColumn.DataPropertyName = "ID"
        Me.IDDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IDDataGridViewTextBoxColumn.Name = "IDDataGridViewTextBoxColumn"
        Me.IDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FFileNameDataGridViewTextBoxColumn
        '
        Me.FFileNameDataGridViewTextBoxColumn.DataPropertyName = "fFileName"
        Me.FFileNameDataGridViewTextBoxColumn.HeaderText = "fFileName"
        Me.FFileNameDataGridViewTextBoxColumn.Name = "FFileNameDataGridViewTextBoxColumn"
        Me.FFileNameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FPathDataGridViewTextBoxColumn
        '
        Me.FPathDataGridViewTextBoxColumn.DataPropertyName = "fPath"
        Me.FPathDataGridViewTextBoxColumn.HeaderText = "fPath"
        Me.FPathDataGridViewTextBoxColumn.Name = "FPathDataGridViewTextBoxColumn"
        Me.FPathDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FSizeDataGridViewTextBoxColumn
        '
        Me.FSizeDataGridViewTextBoxColumn.DataPropertyName = "fSize"
        Me.FSizeDataGridViewTextBoxColumn.HeaderText = "fSize"
        Me.FSizeDataGridViewTextBoxColumn.Name = "FSizeDataGridViewTextBoxColumn"
        Me.FSizeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FCreateDateDataGridViewTextBoxColumn
        '
        Me.FCreateDateDataGridViewTextBoxColumn.DataPropertyName = "fCreateDate"
        Me.FCreateDateDataGridViewTextBoxColumn.HeaderText = "fCreateDate"
        Me.FCreateDateDataGridViewTextBoxColumn.Name = "FCreateDateDataGridViewTextBoxColumn"
        Me.FCreateDateDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FTagsDataGridViewTextBoxColumn
        '
        Me.FTagsDataGridViewTextBoxColumn.DataPropertyName = "fTags"
        Me.FTagsDataGridViewTextBoxColumn.HeaderText = "fTags"
        Me.FTagsDataGridViewTextBoxColumn.Name = "FTagsDataGridViewTextBoxColumn"
        Me.FTagsDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FMarkersDataGridViewTextBoxColumn
        '
        Me.FMarkersDataGridViewTextBoxColumn.DataPropertyName = "fMarkers"
        Me.FMarkersDataGridViewTextBoxColumn.HeaderText = "fMarkers"
        Me.FMarkersDataGridViewTextBoxColumn.Name = "FMarkersDataGridViewTextBoxColumn"
        Me.FMarkersDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FThumbnailDataGridViewImageColumn
        '
        Me.FThumbnailDataGridViewImageColumn.DataPropertyName = "fThumbnail"
        Me.FThumbnailDataGridViewImageColumn.HeaderText = "fThumbnail"
        Me.FThumbnailDataGridViewImageColumn.Name = "FThumbnailDataGridViewImageColumn"
        Me.FThumbnailDataGridViewImageColumn.ReadOnly = True
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(978, 540)
        Me.Controls.Add(Me.DataGridView1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "Form4"
        Me.Text = "Form4"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MVDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MVDataSetBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MarksNoCryptBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents MVDataSetBindingSource As BindingSource
    Friend WithEvents MVDataSet As MVDataSet
    Friend WithEvents MarksNoCryptBindingSource As BindingSource
    Friend WithEvents MarksNoCryptTableAdapter As MVDataSetTableAdapters.MarksNoCryptTableAdapter
    Friend WithEvents IDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FFileNameDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FPathDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FSizeDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FCreateDateDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FTagsDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FMarkersDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents FThumbnailDataGridViewImageColumn As DataGridViewImageColumn
End Class
