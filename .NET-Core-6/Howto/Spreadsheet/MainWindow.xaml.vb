'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'		Sample project that is related to the 'Howto: Use Spreadsheet Formulas in Tables' article inside
'		the 'Windows Presentation Foundation User's Guide'.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Class MainWindow
    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' Window_Loaded Handler
    ' Set the images of the buttons.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Set the images of the buttons.
        SetImages()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_Loaded Handler
    ' Loads the sample document and updates the states of some controls that are related to the TextControl.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Load a sample document.
        Me.m_txTextControl.Load("Files\Cashflow.tx", TXTextControl.StreamType.InternalUnicodeFormat)

        ' Add the supported functions and number formats to the UI dropdowns.
        Dim rstrSupportedFormulaFunctions As String() = Me.m_txTextControl.Tables.SupportedFormulaFunctions
        For Each supportedFormulaFunction In rstrSupportedFormulaFunctions
            Me.m_cmbxFunctions.Items.Add(supportedFormulaFunction)
        Next
        Dim rstrSupportedNumberFormats As String() = Me.m_txTextControl.Tables.SupportedNumberFormats
        For Each supportedNumberFormat In rstrSupportedNumberFormats
            Me.m_cmbxFormats.Items.Add(supportedNumberFormat)
        Next
        Me.m_cmbxFunctions.Text = "SUM"

        ' Set default reference style and enable calculation.
        Me.m_btnEnableCalculation.IsChecked = Me.m_txTextControl.IsFormulaCalculationEnabled
        Me.m_txTextControl.FormulaReferenceStyle = TXTextControl.FormulaReferenceStyle.A1

        ' Check the A1 or R1C1 button.
        If Me.m_txTextControl.FormulaReferenceStyle = TXTextControl.FormulaReferenceStyle.A1 Then
            Me.m_btnA1.IsChecked = True
        Else
            Me.m_btnR1C1.IsChecked = True
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' R1C1_Click Handler
    ' Determines R1C1 as formula reference style and update the UI.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub R1C1_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Determines R1C1 as formula reference style.
        Me.m_txTextControl.FormulaReferenceStyle = TXTextControl.FormulaReferenceStyle.R1C1

        ' Get the current table cell.
        Dim tblCurrentTable As TXTextControl.Table = Me.m_txTextControl.Tables.GetItem()
        If tblCurrentTable IsNot Nothing Then
            Dim tclCurrentTableCell As TXTextControl.TableCell = tblCurrentTable.Cells.GetItem()
            If tclCurrentTableCell IsNot Nothing Then
                ' Update the UI.
                UpdateTableCellSettings(tclCurrentTableCell)
            End If
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' A1_Click Handler
    ' Determines A1 as formula reference style and update the UI.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub A1_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Determines A1 as formula reference style.
        Me.m_txTextControl.FormulaReferenceStyle = TXTextControl.FormulaReferenceStyle.A1

        ' Get the current table cell.
        Dim tblCurrentTable As TXTextControl.Table = Me.m_txTextControl.Tables.GetItem()
        If tblCurrentTable IsNot Nothing Then
            Dim tclCurrentTableCell As TXTextControl.TableCell = tblCurrentTable.Cells.GetItem()
            If tclCurrentTableCell IsNot Nothing Then
                ' Update the UI.
                UpdateTableCellSettings(tclCurrentTableCell)
            End If
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' EnableCalculation_Click Handler
    ' Set a value indicating whether formulas in tables are automatically calculated when the text of an input 
    ' cell is changed.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub EnableCalculation_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.m_txTextControl.IsFormulaCalculationEnabled = Me.m_btnEnableCalculation.IsChecked.Value
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' AddFunction_Click Handler
    ' Add a function to the text box.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub AddFunction_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AddFunction(Me.m_cmbxFunctions.Text)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Formula_KeyUp Handler
    ' Apply the text of the formula text box as new formula to the current table cell when the Return key is 
    ' pressed.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Formula_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.Key = Key.Return Then ' Check whether the Return key is pressed.
            ApplyFormula()
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Accept_Click Handler
    ' Apply a formula to the cell.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Accept_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ApplyFormula()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Remove_Click Handler
    ' Remove the formula from the current table cell and update the UI.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Remove_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Get the current table cell.
        Dim tblCurrentTable As TXTextControl.Table = Me.m_txTextControl.Tables.GetItem()
        If tblCurrentTable IsNot Nothing Then
            Dim tclCurrentTableCell As TXTextControl.TableCell = tblCurrentTable.Cells.GetItem()
            If tclCurrentTableCell IsNot Nothing Then
                ' Remove the formula from the table cell.
                tclCurrentTableCell.Formula = ""
                ' Update the UI.
                UpdateTableCellSettings(tclCurrentTableCell)
            End If
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' TextFormat_Click Handler
    ' Determine that the cell's text is interpreted as text.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextFormat_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        SetCellFormat(TXTextControl.TextType.Standard)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' NumberFormat_Click Handler
    ' Determine that the cell's text is interpreted as a number.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub NumberFormat_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        SetCellFormat(TXTextControl.TextType.Number)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' ApplyNumberFormat_Click Handler
    ' Set the number format for the table cell.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ApplyNumberFormat_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Get the current table cell.
        Dim tblCurrentTable As TXTextControl.Table = Me.m_txTextControl.Tables.GetItem()
        If tblCurrentTable IsNot Nothing Then
            Dim tclCurrentTableCell As TXTextControl.TableCell = tblCurrentTable.Cells.GetItem()
            If tclCurrentTableCell IsNot Nothing Then
                ' Set the number format for the table cell.
                tclCurrentTableCell.CellFormat.NumberFormat = Me.m_cmbxFormats.Text
            End If
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_InputPositionChanged Handler
    ' Enable formula UI when input position is inside a table and a single cell is selected or active
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_InputPositionChanged(ByVal sender As Object, ByVal e As EventArgs)
        ' Check whether the current input position is located inside a table.
        Dim tblCurrentTable As TXTextControl.Table = Me.m_txTextControl.Tables.GetItem()
        If tblCurrentTable Is Nothing Then
            Me.m_tsFormula.IsEnabled = False
            Return
        Else
            ' Get the current table cell.
            Dim tclCurrenTableCell As TXTextControl.TableCell = tblCurrentTable.Cells.GetItem()
            If tclCurrenTableCell IsNot Nothing Then
                Me.m_tsFormula.IsEnabled = True
                UpdateTableCellSettings(tclCurrenTableCell)
            Else
                Me.m_tsFormula.IsEnabled = False
            End If
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' M E T H O D S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' SetImages Method
    ' Sets the images of the buttons.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub SetImages()
        ' m_m_tsFormulaSettings:
        Me.m_imgR1C1.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_EnableR1C1Style.ToString(), Me)
        Me.m_imgA1.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_EnableA1Style.ToString(), Me)
        Me.m_imgEnableCalculation.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_EnableFormulaCalculation.ToString(), Me)

        ' m_m_tsFormula:
        Me.m_imgAddFunction.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_AddFunction.ToString(), Me)
        Me.m_imgAccept.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_AcceptFormula.ToString(), Me)
        Me.m_imgRemove.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_CancelFormulaEditing.ToString(), Me)
        Me.m_imgTextFormat.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_TableCellTextTypeText.ToString(), Me)
        Me.m_imgNumberFormat.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_TableCellTextTypeNumber.ToString(), Me)
        Me.m_imgApplyNumberFormat.Source = TXTextControl.WPF.ResourceProvider.GetSmallIcon(TXTextControl.WPF.RibbonFormulaTab.RibbonItem.TXITEM_AcceptFormula.ToString(), Me)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' AddFunction Method
    ' Add a function to the text box.
    ' Parameters:
    ' 		function	The function to add.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub AddFunction(ByVal [function] As String)
        ' Add the specified function to the text box.
        Me.m_tbxFormula.Text = [function] & "()"
        Me.m_tbxFormula.Select([function].Length + 1, 0)
        Me.m_tbxFormula.Focus()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' ApplyFormula Method
    ' Apply a formula to the current table cell.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ApplyFormula()
        Try
            ' Get the current table cell.
            Dim tblCurrentTable As TXTextControl.Table = Me.m_txTextControl.Tables.GetItem()
            If tblCurrentTable IsNot Nothing Then
                Dim tclCurrentTableCell As TXTextControl.TableCell = tblCurrentTable.Cells.GetItem()
                If tclCurrentTableCell IsNot Nothing Then
                    ' Apply a formula to the current table cell.
                    tclCurrentTableCell.Formula = Me.m_tbxFormula.Text
                End If
            End If
        Catch exc As Exception
            ' Let TXTextControl do the validation.
            MessageBox.Show(exc.Message, "Formula Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' UpdateTableCellSettings Method
    ' Updates the UI based on the specified cell settings.
    ' Parameters:
    '		tableCell	The table cell.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub UpdateTableCellSettings(ByVal tableCell As TXTextControl.TableCell)
        If tableCell IsNot Nothing Then
            ' Read the current formula.
            Me.m_tbxFormula.Text = tableCell.Formula

            ' Set the check state of the reference style.
            Select Case Me.m_txTextControl.FormulaReferenceStyle
                Case TXTextControl.FormulaReferenceStyle.A1
                    Me.m_btnA1.IsChecked = True
                    Me.m_btnR1C1.IsChecked = False
                Case TXTextControl.FormulaReferenceStyle.R1C1
                    Me.m_btnA1.IsChecked = False
                    Me.m_btnR1C1.IsChecked = True
            End Select

            ' Set the check state of the cell format.
            Select Case tableCell.CellFormat.TextType
                Case TXTextControl.TextType.Standard
                    Me.m_btnTextFormat.IsChecked = True
                    Me.m_btnNumberFormat.IsChecked = False
                    Me.m_cmbxFormats.IsEnabled = False
                    Me.m_btnApplyNumberFormat.IsEnabled = False
                    Me.m_cmbxFormats.Text = ""
                Case TXTextControl.TextType.Number
                    Me.m_btnTextFormat.IsChecked = False
                    Me.m_btnNumberFormat.IsChecked = True
                    Me.m_cmbxFormats.IsEnabled = True
                    Me.m_btnApplyNumberFormat.IsEnabled = True
                    Me.m_cmbxFormats.Text = tableCell.CellFormat.NumberFormat
            End Select
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' SetCellFormat Method
    ' Determine that the cell's text is interpreted as text or as a number.
    ' Parameters:
    '		textType	The value how the cell's text is interpreted.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub SetCellFormat(ByVal textType As TXTextControl.TextType)
        ' Get the current table cell.
        Dim tblCurrentTable As TXTextControl.Table = Me.m_txTextControl.Tables.GetItem()
        If tblCurrentTable IsNot Nothing Then
            Dim tclCurrentTableCell As TXTextControl.TableCell = tblCurrentTable.Cells.GetItem()
            If tclCurrentTableCell IsNot Nothing Then
                ' Determine how the cell's text is interpreted.
                tclCurrentTableCell.CellFormat.TextType = textType
                ' Update the UI.
                UpdateTableCellSettings(tclCurrentTableCell)
            End If
        End If
    End Sub
End Class

