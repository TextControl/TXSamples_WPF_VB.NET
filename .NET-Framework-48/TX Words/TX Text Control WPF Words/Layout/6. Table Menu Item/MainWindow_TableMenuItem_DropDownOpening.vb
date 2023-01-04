'-----------------------------------------------------------------------------------------------------------
' MainWindow_TableMenuItem_DropDownOpening.vb File
'
' Description: Provides all SubmenuOpened handlers associated with 'Table' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' Table_SubmenuOpened Handler
        '
        ' Updates the IsEnabled state of 'Table' drop down menu items.
        ' 
        ' Item: 'Table'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Table_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Get current table states
            Dim tblCurrentTable As Table = Me.m_txTextControl.Tables.GetItem()
            Dim bIsTable = tblCurrentTable IsNot Nothing
            Dim bCanTableFormat As Boolean = Me.m_txTextControl.CanTableFormat
            Dim colTableCells = If(bIsTable, tblCurrentTable.Cells, Nothing)
            Dim colTableColumns = If(bIsTable, tblCurrentTable.Columns, Nothing)
            Dim colTableRows = If(bIsTable, tblCurrentTable.Rows, Nothing)


            ' 'Insert'
            Me.m_miTable_Insert.IsEnabled = bCanTableFormat
            Me.m_miTable_Insert_ColumnToTheLeft.IsEnabled = CSharpImpl.Assign(Me.m_miTable_Insert_ColumnToTheRight.IsEnabled, bCanTableFormat AndAlso bIsTable AndAlso colTableColumns.CanAdd)
            Me.m_miTable_Insert_RowAbove.IsEnabled = CSharpImpl.Assign(Me.m_miTable_Insert_RowBelow.IsEnabled, bCanTableFormat AndAlso bIsTable AndAlso colTableRows.CanAdd)

            ' 'Delete'
            Me.m_miTable_Delete.IsEnabled = CSharpImpl.Assign(Me.m_miTable_Delete_Table.IsEnabled, bCanTableFormat AndAlso bIsTable)
            Me.m_miTable_Delete_Cells.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso colTableCells.CanRemove
            Me.m_miTable_Delete_Columns.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso colTableColumns.CanRemove
            Me.m_miTable_Delete_Rows.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso colTableRows.CanRemove

            ' 'Select'
            Me.m_miTable_Select.IsEnabled = CSharpImpl.Assign(Me.m_miTable_Select_Table.IsEnabled, bCanTableFormat AndAlso bIsTable)
            Me.m_miTable_Select_Cell.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso colTableCells.GetItem() IsNot Nothing
            Me.m_miTable_Select_Column.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso colTableColumns.GetItem() IsNot Nothing
            Me.m_miTable_Select_Row.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso colTableRows.GetItem() IsNot Nothing

            ' 'Merge Cells'
            Me.m_miTable_MergeCells.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso tblCurrentTable.CanMergeCells

            ' 'Split Cells'
            Me.m_miTable_SplitCells.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso tblCurrentTable.CanSplitCells

            ' 'Split Table'
            Me.m_miTable_SplitTable.IsEnabled = bCanTableFormat AndAlso bIsTable AndAlso tblCurrentTable.CanSplit

            ' 'Formulas'
            Me.m_miTable_Formulas_A1ReferenceStyle.IsEnabled = CSharpImpl.Assign(Me.m_miTable_Formulas_R1C1ReferenceStyle.IsEnabled, CSharpImpl.Assign(Me.m_miTable_Formulas_AutomaticCalculation.IsEnabled, bCanTableFormat))
            Me.m_miTable_Formulas_EditFormula.IsEnabled = bCanTableFormat AndAlso bIsTable

            ' 'Properties...'
            Me.m_miTable_Properties.IsEnabled = bCanTableFormat AndAlso bIsTable
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Formulas_SubmenuOpened Handler
        '
        ' Updates the checked state of 'Formulas' drop down menu items.
        ' 
        ' Item: 'Formulas'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Formulas_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Set the check states of the 'Formulas' drop down items.
            Dim frsReferenceStyle As FormulaReferenceStyle = Me.m_txTextControl.FormulaReferenceStyle
            Me.m_miTable_Formulas_A1ReferenceStyle.IsChecked = frsReferenceStyle = FormulaReferenceStyle.A1
            Me.m_miTable_Formulas_R1C1ReferenceStyle.IsChecked = frsReferenceStyle = FormulaReferenceStyle.R1C1
            Me.m_miTable_Formulas_AutomaticCalculation.IsChecked = Me.m_txTextControl.IsFormulaCalculationEnabled
        End Sub
    End Class
End Namespace
