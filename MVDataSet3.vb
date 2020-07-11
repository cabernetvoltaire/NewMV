Partial Class MVDataSet
    Partial Public Class MarksNoCryptDataTable
        Private Sub MarksNoCryptDataTable_ColumnChanging(sender As Object, e As DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.fSizeColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class
End Class
