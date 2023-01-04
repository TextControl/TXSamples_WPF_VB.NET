'-----------------------------------------------------------------------------------------------------------
' MainWindow_InsertMenuItem_DropDownOpening.vb File
'
' Description: Provides all SubmenuOpened handlers associated with 'Insert' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' Insert_SubmenuOpened Handler
        '
        ' Updates the IsEnabled state of 'Insert' drop down menu items.
        ' 
        ' Item: 'Insert'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim bCanEdit As Boolean = Me.m_txTextControl.CanEdit
            ' 'File...'
            Me.m_miInsert_File.IsEnabled = bCanEdit

            ' 'Image...'
            Me.m_miInsert_Image.IsEnabled = bCanEdit

            ' 'Text Frame'
            Me.m_miInsert_TextFrame.IsEnabled = bCanEdit

            ' 'Shape'	
            Me.m_miInsert_Shape.IsEnabled = bCanEdit

            ' 'Barcode'
            Me.m_miInsert_Barcode.IsEnabled = bCanEdit

            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                Dim colPages As PageCollection = Me.m_txTextControl.GetPages()
                Dim hfHeaderFooter As HeaderFooter = TryCast(Me.m_txTextControl.TextParts.GetItem(), HeaderFooter)
                Dim pgPage As Page = If(colPages IsNot Nothing, colPages.GetItem(), Nothing)

                ' 'Header'
                Me.m_miInsert_Header_Insert.IsEnabled = colPages IsNot Nothing
                Me.m_miInsert_Header_Remove.IsEnabled = bCanEdit AndAlso pgPage IsNot Nothing AndAlso pgPage.Header IsNot Nothing

                ' 'Footer'
                Me.m_miInsert_Footer_Insert.IsEnabled = colPages IsNot Nothing
                Me.m_miInsert_Footer_Remove.IsEnabled = bCanEdit AndAlso pgPage IsNot Nothing AndAlso pgPage.Footer IsNot Nothing

                ' 'Page Number'
                Me.m_miInsert_PageNumber.IsEnabled = bCanEdit AndAlso hfHeaderFooter IsNot Nothing

                If hfHeaderFooter IsNot Nothing Then
                    Dim pnfPageNumberField As PageNumberField = hfHeaderFooter.PageNumberFields.GetItem()
                    Me.m_miInsert_PageNumber_Insert.IsEnabled = pnfPageNumberField Is Nothing
                    Me.m_miInsert_PageNumber_Delete.IsEnabled = pnfPageNumberField IsNot Nothing
                End If
            End If

            If m_plTXLicense >= VersionInfo.ProductLevel.Enterprise Then
                ' 'Form Fields'
                Me.m_miInsert_FormField_TextFormField.IsEnabled = CSharpImpl.Assign(Me.m_miInsert_FormField_CheckBox.IsEnabled, CSharpImpl.Assign(Me.m_miInsert_FormField_ComboBox.IsEnabled, CSharpImpl.Assign(Me.m_miInsert_FormField_DropDownList.IsEnabled, CSharpImpl.Assign(Me.m_miInsert_FormField_DateFormField.IsEnabled, bCanEdit AndAlso Me.m_txTextControl.FormFields.CanAdd))))
                Me.m_miInsert_FormField_Delete.IsEnabled = bCanEdit AndAlso Me.m_txTextControl.FormFields.GetItem() IsNot Nothing
            End If

            ' 'Symbol'
            Me.m_miInsert_Symbol.IsEnabled = bCanEdit

            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' 'Hyperlink...'
                Me.m_miInsert_Hyperlink.IsEnabled = bCanEdit AndAlso (Me.m_txTextControl.HypertextLinks.CanAdd OrElse Me.m_txTextControl.DocumentLinks.CanAdd)

                ' 'Bookmark...'
                Dim colDocumentTargets As DocumentTargetCollection = Me.m_txTextControl.DocumentTargets
                Me.m_miInsert_Bookmark_Insert.IsEnabled = bCanEdit AndAlso colDocumentTargets.CanAdd
                Me.m_miInsert_Bookmark_Delete.IsEnabled = bCanEdit AndAlso colDocumentTargets.Count <> 0
                Me.m_miInsert_Bookmark.IsEnabled = Me.m_miInsert_Bookmark_Insert.IsEnabled OrElse Me.m_miInsert_Bookmark_Delete.IsEnabled
            End If

            If m_plTXLicense >= VersionInfo.ProductLevel.Enterprise Then
                ' 'Table of Contents'
                Dim bInsideTOC As Boolean = Me.m_txTextControl.TablesOfContents.GetItem() IsNot Nothing
                Me.m_miInsert_TableOfContents_Insert.IsEnabled = bCanEdit AndAlso Not bInsideTOC
                Me.m_miInsert_TableOfContents_Delete.IsEnabled = CSharpImpl.Assign(Me.m_miInsert_TableOfContents_Update.IsEnabled, bCanEdit AndAlso bInsideTOC)
                Me.m_miInsert_TableOfContents.IsEnabled = bCanEdit
            End If

            ' 'Columns'
            Me.m_miInsert_Columns_One.IsEnabled = CSharpImpl.Assign(Me.m_miInsert_Columns_Two.IsEnabled, CSharpImpl.Assign(Me.m_miInsert_Columns_MoreColumns.IsEnabled, bCanEdit))

            ' 'Page Breaks'
            Me.m_miInsert_PageBreaks.IsEnabled = bCanEdit

            ' 'Section Breaks'
            Me.m_miInsert_SectionBreaks.IsEnabled = bCanEdit
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Columns_SubmenuOpened Handler
        '
        ' Updates the checked state of 'Columns' drop down menu items.
        ' 
        ' Item: 'Columns'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Columns_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' Get the number of columns
                Dim secCurrentSection As Section = Me.m_txTextControl.Sections.GetItem()
                Dim iColumns = If(secCurrentSection IsNot Nothing, secCurrentSection.Format.Columns, -1)
                ' Check the items.
                Me.m_miInsert_Columns_One.IsChecked = iColumns = 1
                Me.m_miInsert_Columns_Two.IsChecked = iColumns = 2
                Me.m_miInsert_Columns_MoreColumns.IsChecked = iColumns > 2
            End If
        End Sub
    End Class
End Namespace
