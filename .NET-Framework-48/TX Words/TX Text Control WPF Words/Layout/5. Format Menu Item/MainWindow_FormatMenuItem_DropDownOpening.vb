'-----------------------------------------------------------------------------------------------------------
' MainWindow_FormatMenuItem_DropDownOpening.vb File
'
' Description: Provides all SubmenuOpened handlers associated with 'Format' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' Format_SubmenuOpened Handler
        '
        ' Updates the IsEnabled state of 'Format' drop down menu items.
        ' 
        ' Item: 'Format'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim bCanEdit As Boolean = Me.m_txTextControl.CanEdit
            Dim fbFrame As FrameBase = If(bCanEdit, Me.m_txTextControl.Frames.GetItem(), Nothing)

            ' 'Character...', 
            Me.m_miFormat_Character.IsEnabled = Me.m_txTextControl.CanCharacterFormat

            ' 'Paragraph...'
            Me.m_miFormat_Paragraph.IsEnabled = Me.m_txTextControl.CanParagraphFormat

            ' 'Styles...'
            Me.m_miFormat_Styles.IsEnabled = Me.m_txTextControl.CanStyleFormat

            ' 'Image...'
            Me.m_miFormat_Image.IsEnabled = TypeOf fbFrame Is Image

            ' 'Text Frame...'
            Me.m_miFormat_TextFrame.IsEnabled = TypeOf fbFrame Is TextFrame

            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' 'Shape...'
                Me.m_miFormat_Shape.IsEnabled = EnableShapeItem(fbFrame)
            End If

            ' 'Barcode...'
            Me.m_miFormat_Barcode.IsEnabled = TypeOf fbFrame Is DataVisualization.BarcodeFrame

            ' 'Headers and Footers...'
            Me.m_miFormat_HeadersAndFooters.IsEnabled = bCanEdit

            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' 'Page Number...'
                Dim hfHeaderFooter As HeaderFooter = TryCast(Me.m_txTextControl.TextParts.GetItem(), HeaderFooter)
                Me.m_miFormat_PageNumberField.IsEnabled = bCanEdit AndAlso hfHeaderFooter IsNot Nothing AndAlso hfHeaderFooter.PageNumberFields.GetItem() IsNot Nothing
            End If

            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' 'Hyperlink...'
                Dim colHyperTextLinks As HypertextLinkCollection = Me.m_txTextControl.HypertextLinks
                Me.m_miFormat_Hyperlink.IsEnabled = bCanEdit AndAlso (colHyperTextLinks.GetItem() IsNot Nothing OrElse Me.m_txTextControl.DocumentLinks.GetItem() IsNot Nothing)

                ' 'Bookmark...'
                Dim colDocumentTargets As DocumentTargetCollection = Me.m_txTextControl.DocumentTargets
                Me.m_miFormat_Bookmark.IsEnabled = bCanEdit AndAlso colDocumentTargets.GetItem() IsNot Nothing
            End If

            If m_plTXLicense >= VersionInfo.ProductLevel.Enterprise Then
                ' 'Table of Contents...'
                Dim bInsideTOC As Boolean = Me.m_txTextControl.TablesOfContents.GetItem() IsNot Nothing
                Me.m_miFormat_TableOfContents.IsEnabled = bCanEdit AndAlso bInsideTOC
            End If

            ' 'Columns...'
            Me.m_miFormat_Columns.IsEnabled = bCanEdit

            ' 'Page Borders...'
            Me.m_miFormat_PageBorders.IsEnabled = bCanEdit

            ' 'Page Color...'
            Me.m_miFormat_PageColor.IsEnabled = bCanEdit

            ' 'Tabs...'
            Me.m_miFormat_Tabs.IsEnabled = bCanEdit

            ' 'Language...'
            Me.m_miFormat_Language.IsEnabled = bCanEdit
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_ParagraphStructureLevels_SubmenuOpened Handler
        '
        ' Updates the text of the '[Current Paragraph Style]' item.
        ' 
        ' Item: 'Paragraph Structure Levels'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_ParagraphStructureLevels_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If m_plTXLicense >= VersionInfo.ProductLevel.Enterprise Then
                ' Get the current style name
                Dim strStyleName As String = Me.m_txTextControl.InputFormat.StyleName

                ' Determine current paragraph style
                Dim psCurrentStyle As ParagraphStyle = Me.m_txTextControl.ParagraphStyles.GetItem(strStyleName)

                ' If no paragraph style could be determined, use the default "[Normal]" style.
                If psCurrentStyle Is Nothing Then
                    strStyleName = "[Normal]"
                    psCurrentStyle = Me.m_txTextControl.ParagraphStyles.GetItem(strStyleName)
                End If

                ' Provide the paragraph style by using the item's Tag property.
                Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle.Tag = psCurrentStyle

                ' Display the paragraph style name as item text.
                Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle.Header = String.Format(My.Resources.Item_Format_ParagraphStructureLevels_CurrentParagraphStyle_Text, strStyleName)
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_ParagraphStructureLevels_CurrentParagraphStyle_SubmenuOpened Handler
        '
        ' Updates the checked and IsEnabled state of the '[Current Paragraph Style]' drop down menu items.
        ' 
        ' Item: '[Current Paragraph Style]' of the 'Paragraph Structure Levels' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_ParagraphStructureLevels_CurrentParagraphStyle_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Get the corresponding paragraph style.
            Dim psParagraphStyle As ParagraphStyle = TryCast(Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle.Tag, ParagraphStyle)

            If psParagraphStyle IsNot Nothing Then
                ' Get name and structure level of that style.
                Dim strStyleName = psParagraphStyle.Name
                Dim iStructureLevel = psParagraphStyle.ParagraphFormat.StructureLevel

                ' The strucure levels of the table of contents styles ("TOC_Title" and "TOC_Level") cannot be edited.
                Dim bCanEdit As Boolean = Me.m_txTextControl.CanEdit AndAlso Not (Equals(strStyleName, "TOC_Title") OrElse strStyleName.StartsWith("TOC_Level"))

                ' Step through all structure level drop down items and handle their IsEnabled and Check properties.
                For Each item As MenuItem In Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle.Items
                    item.IsEnabled = bCanEdit
                    item.IsChecked = Integer.Parse(item.Tag.ToString()) = iStructureLevel
                Next
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_ParagraphStructureLevels_AddToParagraph_SubmenuOpened Handler
        '
        ' Updates the checked and IsEnabled state of the 'Add to Paragaph' drop down menu items.
        ' 
        ' Item: 'Add to Paragaph' of the 'Paragraph Structure Levels' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_ParagraphStructureLevels_AddToParagraph_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Check whether the items should be IsEnabled.
            Dim bCanEdit As Boolean = Me.m_txTextControl.CanEdit

            ' Get the current paragraph's structure level.
            Dim iStructureLevel As Integer? = Me.m_txTextControl.InputFormat.StructureLevel
            iStructureLevel = If(iStructureLevel.HasValue, iStructureLevel.Value, -1)

            ' Step through all structure level drop down items and handle their IsEnabled and Check properties.
            For Each item As MenuItem In Me.m_miFormat_ParagraphStructureLevels_AddToParagraph.Items
                item.IsEnabled = bCanEdit
                item.IsChecked = CBool(Integer.Parse(item.Tag.ToString()) = iStructureLevel)
            Next
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_BulletsAndNumbering_SubmenuOpened Handler
        '
        ' Updates the IsEnabled and checked state of 'Bullets and Numbering' drop down menu items.
        ' 
        ' Item: 'Bullets and Numbering'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_BulletsAndNumbering_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Get list format
            Dim lfListFormat As ListFormat = Me.m_txTextControl.Selection.ListFormat

            ' Check list format type
            Dim bIsList = lfListFormat.Type <> ListType.None
            Dim bIsBulleted = lfListFormat.Type = ListType.Bulleted
            Dim bIsStructured = lfListFormat.Type = ListType.Structured

            ' Get number format
            Dim fnNumberFormat = lfListFormat.NumberFormat
            Dim bCanCharacterFormat As Boolean = Me.m_txTextControl.CanCharacterFormat

            ' Set items IsEnabled states
            Me.m_miFormat_BulletsAndNumbering_ArabicNumbers.IsEnabled = CSharpImpl.Assign(Me.m_miFormat_BulletsAndNumbering_CapitalLetters.IsEnabled, CSharpImpl.Assign(Me.m_miFormat_BulletsAndNumbering_Letters.IsEnabled, CSharpImpl.Assign(Me.m_miFormat_BulletsAndNumbering_RomanNumbers.IsEnabled, CSharpImpl.Assign(Me.m_miFormat_BulletsAndNumbering_SmallRomanNumbers.IsEnabled, CSharpImpl.Assign(Me.m_miFormat_BulletsAndNumbering_AsStructuredList.IsEnabled, CSharpImpl.Assign(Me.m_miFormat_BulletsAndNumbering_Bullets.IsEnabled, bCanCharacterFormat))))))                     ' '1, 2, 3'		
            ' 'A, B, C'
            ' 'a, b, c'
            ' 'I, II, III, IV'
            ' 'i, ii, iii, iv'
            ' 'As structured List'
            ' 'Bullets'

            ' 'Increase Level'
            Me.m_miFormat_BulletsAndNumbering_IncreaseLevel.IsEnabled = bIsList AndAlso bCanCharacterFormat

            ' 'Decrease Level'
            Me.m_miFormat_BulletsAndNumbering_DecreaseLevel.IsEnabled = bIsList AndAlso bCanCharacterFormat AndAlso Me.m_txTextControl.Selection.ListFormat.Level >= 2

            ' 'Properties...'
            Me.m_miFormat_BulletsAndNumbering_Properties.IsEnabled = bCanCharacterFormat

            ' Set items IsChecked states
            Me.m_miFormat_BulletsAndNumbering_ArabicNumbers.IsChecked = bIsList AndAlso Not bIsBulleted AndAlso fnNumberFormat = NumberFormat.ArabicNumbers            ' '1, 2, 3'
            Me.m_miFormat_BulletsAndNumbering_CapitalLetters.IsChecked = bIsList AndAlso Not bIsBulleted AndAlso fnNumberFormat = NumberFormat.CapitalLetters          ' 'A, B, C'
            Me.m_miFormat_BulletsAndNumbering_Letters.IsChecked = bIsList AndAlso Not bIsBulleted AndAlso fnNumberFormat = NumberFormat.Letters                        ' 'a, b, c'
            Me.m_miFormat_BulletsAndNumbering_RomanNumbers.IsChecked = bIsList AndAlso Not bIsBulleted AndAlso fnNumberFormat = NumberFormat.RomanNumbers              ' 'I, II, III, IV'
            Me.m_miFormat_BulletsAndNumbering_SmallRomanNumbers.IsChecked = bIsList AndAlso Not bIsBulleted AndAlso fnNumberFormat = NumberFormat.SmallRomanNumbers    ' 'i, ii, iii, iv'
            Me.m_miFormat_BulletsAndNumbering_AsStructuredList.IsChecked = bIsStructured   ' 'As structured List'
            Me.m_miFormat_BulletsAndNumbering_Bullets.IsChecked = bIsBulleted              ' 'Bullets'
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Format_FormFields_SubmenuOpened Handler
        '
        ' Updates the IsEnabled and checked state of 'Form Fields' drop down menu items.
        ' 
        ' Item: 'Form Fields'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Format_FormFields_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If m_plTXLicense >= VersionInfo.ProductLevel.Enterprise Then
                Dim bCanEdit As Boolean = Me.m_txTextControl.CanEdit
                ' 'Form Fields...'
                Dim colFormFields As FormFieldCollection = Me.m_txTextControl.FormFields
                Me.m_miFormat_FormFields_Edit.IsEnabled = bCanEdit AndAlso colFormFields.GetItem() IsNot Nothing

                ' 'Form Validation'
                Me.m_miFormat_FormFields_EnableFormValidation.IsEnabled = bCanEdit AndAlso colFormFields.Count > 0
                Me.m_miFormat_FormFields_EnableFormValidation.IsChecked = CSharpImpl.Assign(Me.m_miFormat_FormFields_ManageConditionalInstructions.IsEnabled, Me.m_txTextControl.IsFormFieldValidationEnabled)
            End If
        End Sub
    End Class
End Namespace
