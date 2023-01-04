'-----------------------------------------------------------------------------------------------------------
' MainWindow_TableMenuItem.vb File
'
' Description: Provides methods to set the layout of the 'Table' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' SetTableItemsTexts Method
        '
        ' Sets the texts of the 'Table' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetTableItemsTexts()
            ' 'Table'
            m_miTable.Header = My.Resources.Item_Table_Text

            ' 'Insert'
            m_miTable_Insert.Header = My.Resources.Item_Table_Insert_Text
            m_miTable_Insert_Table.Header = My.Resources.Item_Table_Insert_Table_Text
            m_miTable_Insert_ColumnToTheLeft.Header = My.Resources.Item_Table_Insert_ColumnToTheLeft_Text
            m_miTable_Insert_ColumnToTheRight.Header = My.Resources.Item_Table_Insert_ColumnToTheRight_Text
            m_miTable_Insert_RowAbove.Header = My.Resources.Item_Table_Insert_RowAbove_Text
            m_miTable_Insert_RowBelow.Header = My.Resources.Item_Table_Insert_RowBelow_Text

            ' 'Delete'
            m_miTable_Delete.Header = My.Resources.Item_Table_Delete_Text
            m_miTable_Delete_Cells.Header = My.Resources.Item_Table_Delete_Cells_Text
            m_miTable_Delete_Columns.Header = My.Resources.Item_Table_Delete_Columns_Text
            m_miTable_Delete_Rows.Header = My.Resources.Item_Table_Delete_Rows_Text
            m_miTable_Delete_Table.Header = My.Resources.Item_Table_Delete_Table_Text

            ' 'Select'
            m_miTable_Select.Header = My.Resources.Item_Table_Select_Text
            m_miTable_Select_Cell.Header = My.Resources.Item_Table_Select_Cell_Text
            m_miTable_Select_Column.Header = My.Resources.Item_Table_Select_Column_Text
            m_miTable_Select_Row.Header = My.Resources.Item_Table_Select_Row_Text
            m_miTable_Select_Table.Header = My.Resources.Item_Table_Select_Table_Text

            ' 'Merge Cells'
            m_miTable_MergeCells.Header = My.Resources.Item_Table_MergeCells_Text

            ' 'Split Cells'
            m_miTable_SplitCells.Header = My.Resources.Item_Table_SplitCells_Text

            ' 'Split Table'
            m_miTable_SplitTable.Header = My.Resources.Item_Table_SplitTable_Text
            m_miTable_SplitTable_Above.Header = My.Resources.Item_Table_SplitTable_Above_Text
            m_miTable_SplitTable_Below.Header = My.Resources.Item_Table_SplitTable_Below_Text

            ' 'Formulas'
            m_miTable_Formulas.Header = My.Resources.Item_Table_Formulas_Text
            m_miTable_Formulas_A1ReferenceStyle.Header = My.Resources.Item_Table_Formulas_A1ReferenceStyle_Text
            m_miTable_Formulas_R1C1ReferenceStyle.Header = My.Resources.Item_Table_Formulas_R1C1ReferenceStyle_Text
            m_miTable_Formulas_EditFormula.Header = My.Resources.Item_Table_Formulas_EditFormula_Text
            m_miTable_Formulas_AutomaticCalculation.Header = My.Resources.Item_Table_Formulas_AutomaticCalculation_Text

            ' 'Properties...'
            m_miTable_Properties.Header = My.Resources.Item_Table_Properties_Text
        End Sub
        '-----------------------------------------------------------------------------------------------------------
        ' SetTableItemsImages Method
        '
        ' Sets the images of the 'Table' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetTableItemsImages()
            ' 'Insert'
            Me.SetItemImage(Me.m_miTable_Insert, RibbonInsertTab.RibbonItem.TXITEM_InsertTable.ToString())
            Me.SetItemImage(Me.m_miTable_Insert_Table, RibbonInsertTab.RibbonItem.TXITEM_InsertTable.ToString())
            Me.SetItemImage(Me.m_miTable_Insert_ColumnToTheLeft, RibbonTableLayoutTab.RibbonItem.TXITEM_InsertTableColLeft.ToString())
            Me.SetItemImage(Me.m_miTable_Insert_ColumnToTheRight, RibbonTableLayoutTab.RibbonItem.TXITEM_InsertTableColRight.ToString())
            Me.SetItemImage(Me.m_miTable_Insert_RowAbove, RibbonTableLayoutTab.RibbonItem.TXITEM_InsertTableRowAbove.ToString())
            Me.SetItemImage(Me.m_miTable_Insert_RowBelow, RibbonTableLayoutTab.RibbonItem.TXITEM_InsertTableRowBelow.ToString())

            ' 'Delete'
            Me.SetItemImage(Me.m_miTable_Delete, RibbonTableLayoutTab.RibbonItem.TXITEM_DeleteTable.ToString())
            Me.SetItemImage(Me.m_miTable_Delete_Table, RibbonTableLayoutTab.RibbonItem.TXITEM_DeleteTable.ToString())
            Me.SetItemImage(Me.m_miTable_Delete_Columns, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_DeleteTableCol.ToString())
            Me.SetItemImage(Me.m_miTable_Delete_Rows, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_DeleteTableRow.ToString())
            Me.SetItemImage(Me.m_miTable_Delete_Cells, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_DeleteTableCell.ToString())

            ' 'Select'
            Me.SetItemImage(Me.m_miTable_Select, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_SelectTableRow.ToString())
            Me.SetItemImage(Me.m_miTable_Select_Table, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_SelectTableAll.ToString())
            Me.SetItemImage(Me.m_miTable_Select_Column, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_SelectTableCol.ToString())
            Me.SetItemImage(Me.m_miTable_Select_Row, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_SelectTableRow.ToString())
            Me.SetItemImage(Me.m_miTable_Select_Cell, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_SelectTableCell.ToString())

            ' 'Merge Cells'
            Me.SetItemImage(Me.m_miTable_MergeCells, RibbonTableLayoutTab.RibbonItem.TXITEM_MergeTableCells.ToString())

            ' 'Split Cells'
            Me.SetItemImage(Me.m_miTable_SplitCells, RibbonTableLayoutTab.RibbonItem.TXITEM_SplitTableCells.ToString())

            ' 'Split Table'
            Me.SetItemImage(Me.m_miTable_SplitTable, RibbonTableLayoutTab.RibbonItem.TXITEM_SplitTable.ToString())
            Me.SetItemImage(Me.m_miTable_SplitTable_Above, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_SplitTableAbove.ToString())
            Me.SetItemImage(Me.m_miTable_SplitTable_Below, RibbonTableLayoutTab.RibbonDropDownItem.TXITEM_SplitTableBelow.ToString())

            ' 'Formulas'
            Me.SetItemImage(Me.m_miTable_Formulas, RibbonFormulaTab.RibbonItem.TXITEM_AddFunction.ToString())
            Me.SetItemImage(Me.m_miTable_Formulas_A1ReferenceStyle, RibbonFormulaTab.RibbonItem.TXITEM_EnableA1Style.ToString())
            Me.SetItemImage(Me.m_miTable_Formulas_R1C1ReferenceStyle, RibbonFormulaTab.RibbonItem.TXITEM_EnableR1C1Style.ToString())
            Me.SetItemImage(Me.m_miTable_Formulas_EditFormula, "TXITEM_FormulaEditing")
            Me.SetItemImage(Me.m_miTable_Formulas_AutomaticCalculation, RibbonFormulaTab.RibbonItem.TXITEM_EnableFormulaCalculation.ToString())

            ' 'Properties...'
            Me.SetItemImage(Me.m_miTable_Properties, RibbonInsertTab.RibbonDropDownItem.TXITEM_InsertTableDialog.ToString())
        End Sub
    End Class
End Namespace
