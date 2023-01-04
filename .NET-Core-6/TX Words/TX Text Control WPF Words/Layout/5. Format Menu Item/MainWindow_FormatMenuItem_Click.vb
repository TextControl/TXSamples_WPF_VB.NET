'-----------------------------------------------------------------------------------------------------------
' MainWindow_FormatMenuItem_Click.vb File
'
' Description: Provides all Click handlers associated with 'Format' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.WPF
Imports TXTextControl.WPF.Drawing

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' Format_Character_Click Handler
        '
        ' Invokes the built-in dialog box for setting fonts and character attributes.
        ' 
        ' Item: 'Character...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Character_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.FontDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Paragraph_Click Handler
        '
        ' Invokes the built-in dialog box for setting the formatting attributes of a paragraph.
        ' 
        ' Item: 'Paragraph...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Paragraph_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.ParagraphFormatDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_ParagraphStructureLevels_CurrentParagraphStyle_Level_Click Handler
        '
        ' Sets the structure level of the specified paragraph style (see ParagraphStructureLevels_SubmenuOpened 
        ' handler). The level is determined by the Tag property value of the clicked item.
        ' 
        ' Item: Each item of the 'Paragraph Style [Current Praragraph Style]' drop down menu ('Paragraph Style 
        '		 [Current Praragraph Style]' is an item of the 'Paragraph Structure Levels' drop down menu)
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_ParagraphStructureLevels_CurrentParagraphStyle_Level_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Enforce license exception if paragraph styles are not supported by the current product level.
            If m_plTXLicense < VersionInfo.ProductLevel.Enterprise Then
                Me.m_txTextControl.ParagraphStyles.GetItem("[Normal]")
            End If
            ' Get the paragraph style.
            Dim psParagraphStyle As ParagraphStyle = TryCast(Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle.Tag, ParagraphStyle)
            ' Set structure level to the paragraph style.
            psParagraphStyle.ParagraphFormat.StructureLevel = Integer.Parse(TryCast(e.Source, MenuItem).Tag.ToString())
            ' Apply the set level to the document.
            psParagraphStyle.Apply()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_ParagraphStructureLevels_AddToParagraph_Level_Click Handler
        '
        ' Sets the structure level to the all selected paragraphs.
        ' 
        ' Item: Each item of the 'Add to Paragraph' drop down menu ('Add to Paragraph' is an item of the 
        '		 'Paragraph Structure Levels' drop down menu)
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_ParagraphStructureLevels_AddToParagraph_Level_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.InputFormat.StructureLevel = Integer.Parse(TryCast(e.Source, MenuItem).Tag.ToString())
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Styles_Click Handler
        '
        ' Invokes the built-in dialog box for creating, deleting and modifying formatting styles.
        ' 
        ' Item: 'Styles...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Styles_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.FormattingStylesDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_ArabicNumbers_Click Handler
        '
        ' Determines whether there is a numbered or structured list with arabic numbers at the current input position. 
        ' 
        ' Item: '1, 2, 3' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_ArabicNumbers_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.SetNumberedList(Me.m_miFormat_BulletsAndNumbering_ArabicNumbers.IsChecked, NumberFormat.ArabicNumbers)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_CapitalLetters_Click Handler
        '
        ' Determines whether there is a numbered or structured list with capital letters at the current input position. 
        ' 
        ' Item: 'A, B, C' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_CapitalLetters_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.SetNumberedList(Me.m_miFormat_BulletsAndNumbering_CapitalLetters.IsChecked, NumberFormat.CapitalLetters)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_Letters_Click Handler
        '
        ' Determines whether there is a numbered or structured list with small letters at the current input position. 
        ' 
        ' Item: 'a, b, c' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_Letters_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.SetNumberedList(Me.m_miFormat_BulletsAndNumbering_Letters.IsChecked, NumberFormat.Letters)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_RomanNumbers_Click Handler
        '
        ' Determines whether there is a numbered or structured list with capital roman numbers at the current input 
        ' position. 
        ' 
        ' Item: 'I, II, III, IV' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_RomanNumbers_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.SetNumberedList(Me.m_miFormat_BulletsAndNumbering_RomanNumbers.IsChecked, NumberFormat.RomanNumbers)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_SmallRomanNumbers_Click Handler
        '
        ' Determines whether there is a numbered or structured list with small roman numbers at the current input 
        ' position. 
        ' 
        ' Item: 'i, ii, iii, iv' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_SmallRomanNumbers_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.SetNumberedList(Me.m_miFormat_BulletsAndNumbering_SmallRomanNumbers.IsChecked, NumberFormat.SmallRomanNumbers)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_AsStructuredList_Click Handler
        '
        ' Determines whether the current list is a numbered or a structured list.
        ' 
        ' Item: 'As structured List' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_AsStructuredList_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Check whether a numbered list is currently set at the current input position.
            Dim bIsNumberedList As Boolean? = Me.m_txTextControl.InputFormat.NumberedList

            If bIsNumberedList.HasValue AndAlso bIsNumberedList.Value Then
                ' In this case change that list to a structured list.
                Me.m_txTextControl.InputFormat.StructuredList = True
                Return
            End If
            ' Check whether a structured list is currently set at the current input position.
            Dim bIsStructuredList As Boolean? = Me.m_txTextControl.InputFormat.StructuredList

            If bIsStructuredList.HasValue AndAlso bIsStructuredList.Value Then
                ' In this case change that list to a numbered list.
                Me.m_txTextControl.InputFormat.NumberedList = True
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_Bullets_Click Handler
        '
        ' Determines whether the current list is a bulleted list.
        ' 
        ' Item: 'Bullets' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_Bullets_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.InputFormat.BulletedList = Me.m_miFormat_BulletsAndNumbering_Bullets.IsChecked
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_IncreaseLevel_Click Handler
        '
        ' Increases the list format level and the indent.
        ' 
        ' Item: 'Increase Level' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_IncreaseLevel_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Selection.ListFormat.Level += 1
            Me.m_txTextControl.Selection.IncreaseIndent()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_DecreaseLevel_Click Handler
        '
        ' Decreases the list format level and the indent.
        ' 
        ' Item: 'Decrease Level' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_DecreaseLevel_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Selection.ListFormat.Level -= 1
            Me.m_txTextControl.Selection.DecreaseIndent()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_Properties_Click Handler
        '
        ' Invokes the built-in dialog box for setting formatting attributes of bulleted and numbered lists.
        ' 
        ' Item: 'Properties...' of the 'Bullets and Numbering' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_Properties_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.ListFormatDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Image_Click Handler
        '
        ' Invokes the built-in dialog box for setting attributes of the selected image. 
        ' 
        ' Item: 'Image...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Image_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.ImageAttributesDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_TextFrame_Click Handler
        '
        ' Invokes the built-in dialog box for setting attributes of the selected text frame.
        ' 
        ' Item: 'Text Frame...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_TextFrame_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TextFrameAttributesDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Shape_Click Handler
        '
        ' Invokes the built-in dialog box for alter the layout settings, the size and the text distances of the 
        ' selected drawing frame. If the drawing frame is activated, a dialog is opened to format the selected shapes 
        ' of the drawing canvas.
        ' 
        ' Item: 'Shape...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Shape_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Check whether a drawing frame is activated.
            Dim dfDrawingFrame As DataVisualization.DrawingFrame = Me.m_txTextControl.Drawings.GetActivatedItem()

            If dfDrawingFrame IsNot Nothing Then
                ' In that case open the format shapes dialog.
                Dim txdDrawingControl As TXDrawingControl = TryCast(dfDrawingFrame.Drawing, TXDrawingControl)
                txdDrawingControl.FormatShapesDialog()
            Else
                ' Otherwise the layout dialog for the selected drawing frame is opened.
                Me.m_txTextControl.DrawingLayoutDialog()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Frame_Barcode_Click Handler
        '
        ' Invokes the built-in dialog box for alter the layout settings, the size and the text distances of the 
        ' selected barcode. 
        ' 
        ' Item: 'Barcode...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Frame_Barcode_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.BarcodeLayoutDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Frame_HeadersAndFooters_Click Handler
        '
        ' Invokes the second tab of the built-in section dialog for specifying headers and footers.
        ' 
        ' Item: 'Headers and Footers...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Frame_HeadersAndFooters_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.SectionFormatDialog(1)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_PageNumberField_Click Handler
        '
        ' Opens a dialog box to alter the formatting and numbering attributes of the page number field.
        ' 
        ' Item: 'Page Number...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_PageNumberField_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim hfHeaderFooter As HeaderFooter = TryCast(Me.m_txTextControl.TextParts.GetItem(), HeaderFooter)
            Dim pnfCurrentPageNumberField As PageNumberField = hfHeaderFooter.PageNumberFields.GetItem()
            pnfCurrentPageNumberField.PageNumberDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_FormFields_Edit_Click Handler
        '
        ' Opens a dialog box to format the current form field.
        ' 
        ' Item: 'Edit' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_FormFields_Edit_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim ffFormField As FormField = m_txTextControl.FormFields.GetItem()
            ffFormField.FormFieldDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_FormFields_EnableFormValidation_Click Handler
        '
        ' Sets a value indicating whether Conditional Instructions are applied to form fields when the EditMode 
        ' property is set to EditMode.ReadAndSelect and TextControl.DocumentPermissions.ReadOnly to true. 
        ' 
        ' Item: 'Enable Form Validation' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_FormFields_EnableFormValidation_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.IsFormFieldValidationEnabled = Me.m_miFormat_FormFields_EnableFormValidation.IsChecked
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_FormFields_ManageConditionalInstructions_Click Handler
        '
        ' Opens a dialog box to add, edit or delete Conditional Instructions inside the document.
        ' 
        ' Item: 'Manage Conditional Instructions...' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_FormFields_ManageConditionalInstructions_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.ManageConditionalInstructionsDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Hyperlink_Click Handler
        '
        ' Opens a built-in dialog box for editing the HypertextLink or DocumentLink at the current text input 
        ' position.
        ' 
        ' Item: 'Hyperlink...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Hyperlink_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim dlgHyperlink As HyperlinkDialog = New HyperlinkDialog(Me.m_txTextControl) With {
                .Owner = Me
            }
            dlgHyperlink.ShowDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Bookmark_Click Handler
        '
        ' Opens a built-in dialog box for editing the DocumentTarget at the current text input position.
        ' 
        ' Item: 'Bookmark...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Bookmark_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim dlgEditBookmark As BookmarkDialog = New BookmarkDialog(Me.m_txTextControl, Me.m_txTextControl.DocumentTargets.GetItem()) With {
                .Owner = Me
            }
            dlgEditBookmark.ShowDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_TableOfContents_Click Handler
        '
        ' Invokes the built-in dialog box for formatting a table of contents. 
        ' 
        ' Item: 'Table of Contents...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_TableOfContents_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TableOfContentsDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Columns_Click Handler
        '
        ' Invokes the third tab of the built-in section dialog for specifying the attributes of the page columns.
        ' 
        ' Item: 'Columns...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Columns_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.SectionFormatDialog(2)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_PageBorders_Click Handler
        '
        ' Invokes the fourth tab of the built-in section dialog for specifying page borders
        ' 
        ' Item: 'Page Borders...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_PageBorders_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.SectionFormatDialog(3)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_PageColor_Click Handler
        '
        ' Invokes the built-in dialog box for setting the page color.
        ' 
        ' Item: 'Page Color...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_PageColor_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.PageColorDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Tabs_Click Handler
        '
        ' Invokes the built-in dialog box for setting tabs.
        ' 
        ' Item: 'Tabs...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Tabs_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TabDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_Language_Click Handler
        '
        ' Invokes the built-in dialog box for setting the language of the selected text.
        ' 
        ' Item: 'Language...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_Language_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.LanguageDialog()
        End Sub
    End Class
End Namespace
