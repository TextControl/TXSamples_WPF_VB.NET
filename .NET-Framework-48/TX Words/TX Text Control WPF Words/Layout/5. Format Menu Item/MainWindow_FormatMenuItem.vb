'-----------------------------------------------------------------------------------------------------------
' MainWindow_FormatMenuItem.vb File
'
' Description: Provides methods to set the layout of the 'Format' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetFormatItemsTexts Method
        '
        ' Sets the texts of the 'Format' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetFormatItemsTexts()
            ' 'Format'
            m_miFormat.Header = My.Resources.Item_Format_Text

            ' 'Character...'
            m_miFormat_Character.Header = My.Resources.Item_Format_Character_Text

            ' 'Paragraph...'
            m_miFormat_Paragraph.Header = My.Resources.Item_Format_Paragraph_Text

            ' 'Styles...'
            m_miFormat_Styles.Header = My.Resources.Item_Format_Styles_Text

            ' 'Paragraph Structure Levels'
            m_miFormat_ParagraphStructureLevels.Header = My.Resources.Item_Format_ParagraphStructureLevels_Text
            m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle_BodyText.Header = My.Resources.Item_Format_ParagraphStructureLevels_CurrentParagraphStyle_BodyText_Text
            ' Set texts of the 'Paragraph Style: [Current Paragraph Style]' item's 'Level' drop down items.
            For i As Integer = 1 To Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle.Items.Count - 1
                ' Get item.
                Dim miLevel As MenuItem = TryCast(Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle.Items(i), MenuItem)
                ' Create accelerator string.
                Dim strLevel = If(i < 10, "_" & i, "1_0")
                ' Set text.
                SetItemText(miLevel, My.Resources.Item_Format_ParagraphStructureLevels_CurrentParagraphStyle_Level_Text, strLevel)
            Next

            m_miFormat_ParagraphStructureLevels_AddToParagraph.Header = My.Resources.Item_Format_ParagraphStructureLevels_AddToParagraph_Text
            m_miFormat_ParagraphStructureLevels_AddToParagraph_BodyText.Header = My.Resources.Item_Format_ParagraphStructureLevels_AddToParagraph_BodyText_Text
            ' Set texts of the 'Add to Paragraph' item's 'Level' drop down items.
            For i As Integer = 1 To Me.m_miFormat_ParagraphStructureLevels_AddToParagraph.Items.Count - 1
                ' Get item.
                Dim miLevel As MenuItem = TryCast(Me.m_miFormat_ParagraphStructureLevels_AddToParagraph.Items(i), MenuItem)
                ' Create accelerator string.
                Dim strLevel = If(i < 10, "_" & i, "1_0")
                ' Set text.
                SetItemText(miLevel, My.Resources.Item_Format_ParagraphStructureLevels_AddToParagraph_Level_Text, strLevel)
            Next

            ' 'Bullets and Numbering'
            m_miFormat_BulletsAndNumbering.Header = My.Resources.Item_Format_BulletsAndNumbering_Text
            m_miFormat_BulletsAndNumbering_ArabicNumbers.Header = My.Resources.Item_Format_BulletsAndNumbering_ArabicNumbers_Text
            m_miFormat_BulletsAndNumbering_CapitalLetters.Header = My.Resources.Item_Format_BulletsAndNumbering_CapitalLetters_Text
            m_miFormat_BulletsAndNumbering_Letters.Header = My.Resources.Item_Format_BulletsAndNumbering_Letters_Text
            m_miFormat_BulletsAndNumbering_RomanNumbers.Header = My.Resources.Item_Format_BulletsAndNumbering_RomanNumbers_Text
            m_miFormat_BulletsAndNumbering_SmallRomanNumbers.Header = My.Resources.Item_Format_BulletsAndNumbering_SmallRomanNumbers_Text
            m_miFormat_BulletsAndNumbering_AsStructuredList.Header = My.Resources.Item_Format_BulletsAndNumbering_AsStructuredList_Text
            m_miFormat_BulletsAndNumbering_Bullets.Header = My.Resources.Item_Format_BulletsAndNumbering_Bullets_Text
            m_miFormat_BulletsAndNumbering_IncreaseLevel.Header = My.Resources.Item_Format_BulletsAndNumbering_IncreaseLevel_Text
            m_miFormat_BulletsAndNumbering_DecreaseLevel.Header = My.Resources.Item_Format_BulletsAndNumbering_DecreaseLevel_Text
            m_miFormat_BulletsAndNumbering_Properties.Header = My.Resources.Item_Format_BulletsAndNumbering_Properties_Text

            ' 'Image...'
            m_miFormat_Image.Header = My.Resources.Item_Format_Image_Text

            ' 'Text Frame...'
            m_miFormat_TextFrame.Header = My.Resources.Item_Format_TextFrame_Text

            ' 'Shape...'
            m_miFormat_Shape.Header = My.Resources.Item_Format_Shape_Text

            ' 'Barcode...'
            m_miFormat_Barcode.Header = My.Resources.Item_Format_Barcode_Text

            ' 'Headers and Footers...'
            m_miFormat_HeadersAndFooters.Header = My.Resources.Item_Format_HeadersAndFooters_Text

            ' 'Page Number...'
            m_miFormat_PageNumberField.Header = My.Resources.Item_Format_PageNumberField_Text

            ' 'Form Fields'
            m_miFormat_FormFields.Header = My.Resources.Item_Format_FormFields_Text
            m_miFormat_FormFields_Edit.Header = My.Resources.Item_Format_FormFields_Edit_Text
            m_miFormat_FormFields_EnableFormValidation.Header = My.Resources.Item_Format_FormFields_EnableFormValidation_Text
            m_miFormat_FormFields_ManageConditionalInstructions.Header = My.Resources.Item_Format_FormFields_ManageConditionalInstructions_Text

            ' 'Hyperlink...'
            m_miFormat_Hyperlink.Header = My.Resources.Item_Format_Hyperlink_Text

            ' 'Bookmark...'
            m_miFormat_Bookmark.Header = My.Resources.Item_Format_Bookmark_Text

            ' 'Table of Contents...'
            m_miFormat_TableOfContents.Header = My.Resources.Item_Format_TableOfContents_Text

            ' 'Columns...'
            m_miFormat_Columns.Header = My.Resources.Item_Format_Columns_Text

            ' 'Page Borders...'
            m_miFormat_PageBorders.Header = My.Resources.Item_Format_PageBorders_Text

            ' 'Page Color...'
            m_miFormat_PageColor.Header = My.Resources.Item_Format_PageColor_Text

            ' 'Tabs...'
            m_miFormat_Tabs.Header = My.Resources.Item_Format_Tabs_Text

            ' 'Language...'
            m_miFormat_Language.Header = My.Resources.Item_Format_Language_Text
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SetFormatItemsImages Method
        '
        ' Sets the images of the 'Format' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetFormatItemsImages()
            ' 'Character...'
            Me.SetItemImage(Me.m_miFormat_Character, RibbonFormattingTab.RibbonItem.TXITEM_ChangeCase.ToString())

            ' 'Paragraph...'
            Me.SetItemImage(Me.m_miFormat_Paragraph, RibbonFormattingTab.RibbonItem.TXITEM_ControlChars.ToString())

            ' 'Styles...'
            Me.SetItemImage(Me.m_miFormat_Styles, RibbonFormattingTab.RibbonItem.TXITEM_StyleName.ToString())

            ' 'Paragraph Structure Levels'
            Me.SetItemImage(Me.m_miFormat_ParagraphStructureLevels, RibbonReferencesTab.RibbonItem.TXITEM_TOCMinimumStructureLevel.ToString())
            Me.SetItemImage(Me.m_miFormat_ParagraphStructureLevels_CurrentParagraphStyle, RibbonFormattingTab.RibbonItem.TXITEM_StyleName.ToString())
            Me.SetItemImage(Me.m_miFormat_ParagraphStructureLevels_AddToParagraph, RibbonFormattingTab.RibbonItem.TXITEM_ControlChars.ToString())

            ' 'Bullets and Numbering'
            Me.SetItemImage(Me.m_miFormat_BulletsAndNumbering, RibbonFormattingTab.RibbonItem.TXITEM_NumberedList.ToString())
            Me.SetItemImage(Me.m_miFormat_BulletsAndNumbering_AsStructuredList, RibbonFormattingTab.RibbonItem.TXITEM_StructuredList.ToString())
            Me.SetItemImage(Me.m_miFormat_BulletsAndNumbering_Bullets, RibbonFormattingTab.RibbonItem.TXITEM_BulletedList.ToString())
            Me.SetItemImage(Me.m_miFormat_BulletsAndNumbering_IncreaseLevel, RibbonFormattingTab.RibbonItem.TXITEM_IncreaseIndent.ToString())
            Me.SetItemImage(Me.m_miFormat_BulletsAndNumbering_DecreaseLevel, RibbonFormattingTab.RibbonItem.TXITEM_DecreaseIndent.ToString())
            Me.SetItemImage(Me.m_miFormat_BulletsAndNumbering_Properties, RibbonFormattingTab.RibbonItem.TXITEM_NumberedList.ToString())

            ' 'Image...'
            Me.SetItemImage(Me.m_miFormat_Image, RibbonInsertTab.RibbonItem.TXITEM_InsertImage.ToString())

            ' 'Text Frame...'
            Me.SetItemImage(Me.m_miFormat_TextFrame, RibbonInsertTab.RibbonItem.TXITEM_InsertTextFrame.ToString())

            ' 'Shape...'
            Me.SetItemImage(Me.m_miFormat_Shape, RibbonInsertTab.RibbonItem.TXITEM_InsertShape.ToString())

            ' 'Barcode...'
            Me.SetItemImage(Me.m_miFormat_Barcode, RibbonInsertTab.RibbonItem.TXITEM_InsertBarcode.ToString())

            ' 'Headers and Footers...'
            Me.SetItemImage(Me.m_miFormat_HeadersAndFooters, RibbonInsertTab.RibbonItem.TXITEM_InsertHeader.ToString())

            ' 'Page Number...'
            Me.SetItemImage(Me.m_miFormat_PageNumberField, RibbonInsertTab.RibbonItem.TXITEM_InsertPageNumber.ToString())

            ' 'Form Fields'
            Me.SetItemImage(Me.m_miFormat_FormFields, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertComboBoxField.ToString())
            Me.SetItemImage(Me.m_miFormat_FormFields_Edit, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertComboBoxField.ToString())
            Me.SetItemImage(Me.m_miFormat_FormFields_EnableFormValidation, RibbonFormFieldsTab.RibbonItem.TXITEM_EnableFormValidation.ToString())
            Me.SetItemImage(Me.m_miFormat_FormFields_ManageConditionalInstructions, RibbonFormFieldsTab.RibbonItem.TXITEM_ManageConditionalInstructions.ToString())

            ' 'Hyperlink...'
            Me.SetItemImage(Me.m_miFormat_Hyperlink, RibbonInsertTab.RibbonItem.TXITEM_InsertHyperlink.ToString())

            ' 'Bookmark...'
            Me.SetItemImage(Me.m_miFormat_Bookmark, RibbonInsertTab.RibbonItem.TXITEM_InsertBookmark.ToString())

            ' 'Table of Contents...'
            Me.SetItemImage(Me.m_miFormat_TableOfContents, RibbonReferencesTab.RibbonItem.TXITEM_ModifyTableOfContents.ToString())

            ' 'Columns...'
            Me.SetItemImage(Me.m_miFormat_Columns, RibbonPageLayoutTab.RibbonItem.TXITEM_Columns.ToString())

            ' 'Page Borders...'
            Me.SetItemImage(Me.m_miFormat_PageBorders, RibbonPageLayoutTab.RibbonItem.TXITEM_PageBorders.ToString())

            ' 'Page Color...'
            Me.SetItemImage(Me.m_miFormat_PageColor, RibbonPageLayoutTab.RibbonItem.TXITEM_PageColor.ToString())

            ' 'Tabs...'
            Me.SetItemImage(Me.m_miFormat_Tabs, RibbonFormattingTab.RibbonItem.TXITEM_EditTabs.ToString())

            ' 'Language...'
            Me.SetItemImage(Me.m_miFormat_Language, RibbonProofingTab.RibbonItem.TXITEM_SetLanguage.ToString())
        End Sub
    End Class
End Namespace
