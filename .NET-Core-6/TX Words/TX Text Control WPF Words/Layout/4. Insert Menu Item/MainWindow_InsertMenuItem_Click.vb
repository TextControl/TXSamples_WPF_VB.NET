'-----------------------------------------------------------------------------------------------------------
' MainWindow_InsertMenuItem_Click.vb File
'
' Description: Provides all Click handlers associated with 'Insert' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.DataVisualization
Imports TXTextControl.WPF
Imports TXTextControl.WPF.Barcode
Imports TXTextControl.WPF.Drawing

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' Insert_File_Click Handler
        '
        ' Opens a dialog to exchange the currently selected text with a specified file.
        ' 
        ' Item: 'File...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_File_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Create load settings.
            Dim lsLoadSettings As LoadSettings = New LoadSettings With {
                .ApplicationFieldFormat = ApplicationFieldFormat.MSWordTXFormFields,
                .LoadSubTextParts = True,
                .DocumentPartName = String.Empty
            }
            ' Open the dialog to chose a file that exchanges the currently selected text.
            Me.m_txTextControl.Selection.Load(StreamType.All, lsLoadSettings)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Image_Click Handler
        '
        ' Opens a dialog to insert an image where the text flows around that image and empty areas at the left and 
        ' right side are filled.
        ' 
        ' Item: 'Image...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Image_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Images.Add(New Image(), TXTextControl.HorizontalAlignment.Left, -1, ImageInsertionMode.DisplaceText)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_TextFrame_Click Handler
        '
        ' Inserts a text frame by sizing its bounds with the mouse. The text frame is anchored to a paragraph and 
        ' moves with the text. The text flows around the text frame and empty areas at the left and right side are 
        ' filled.
        ' 
        ' Item: 'Text Frame'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_TextFrame_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TextFrames.Add(New TextFrame(New System.Drawing.Size(2880, 2880)), TextFrameInsertionMode.DisplaceText Or TextFrameInsertionMode.MoveWithText)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Shape_ShapeCategory_MenuItem_Click Handler
        '
        ' Inserts a shape as a DrawingFrame into the TextControl or, if a drawing frame is activated, into it. If the
        ' shape is inserted as a DrawingFrame, it is anchored to a paragraph and moves with the text. The text flows 
        ' around the drawing frame and empty areas at the left and right side are filled.
        '  
        ' Item: Each item of the 'Shape' drop down menu's category items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Shape_ShapeCategory_MenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            AddShape(CType(TryCast(e.Source, MenuItem).Tag, Drawing.ShapeType))
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Shape_DrawingCanvas_Click Handler
        '
        ' Inserts an activated DrawingFrame that represents a drawing canvas into the TextControl. The text flows  
        ' around that image and empty areas at the left and right side are filled.
        '  
        ' Item: 'Drawing Canvas' of the 'Shape' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Shape_DrawingCanvas_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Add a DrawingFrame that represents a drawing canvas into the TextControl.
            Dim drawing As TXDrawingControl = New TXDrawingControl(7000, 4000)
            Dim dfDrawingFrame As DrawingFrame = New DrawingFrame(drawing)
            Me.m_txTextControl.Drawings.Add(dfDrawingFrame, TXTextControl.HorizontalAlignment.Left, -1, FrameInsertionMode.DisplaceText)

            ' Activate the DrawingFrame.
            dfDrawingFrame.Activate()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Barcode_MenuItem_Click Handler
        '
        ' Inserts a BarcodeFrame that represents the barcode type that is represented by the clicked item into the 
        ' TextControl. It is anchored to a paragraph and moves with the text. The text flows around the barcode 
        ' frame and empty areas at the left and right side are filled.
        '  
        ' Item: Each item of the 'Barcode' drop down items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Barcode_MenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim txbBarcodeControl As TXBarcodeControl = New TXBarcodeControl With {
                .Width = 240,
                .Height = 180,
                .BarcodeType = CType(TryCast(e.Source, MenuItem).Tag, Barcode.BarcodeType)
            }
            Dim bfBarcodeFrame As BarcodeFrame = New BarcodeFrame(txbBarcodeControl)
            Me.m_txTextControl.Barcodes.Add(bfBarcodeFrame, TXTextControl.HorizontalAlignment.Left, -1, FrameInsertionMode.DisplaceText Or FrameInsertionMode.MoveWithText)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Header_Insert_Click Handler
        '
        ' Insert a header to the TextControl (or activates the header if it already exists).
        '  
        ' Item: 'Insert' of the 'Header' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Header_Insert_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            InsertHeaderFooter(HeaderFooterType.Header)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Header_Remove_Click Handler
        '
        ' Removes the header from the TextControl.
        '  
        ' Item: 'Remove' of the 'Header' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Header_Remove_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim hfHeader As HeaderFooter = Me.m_txTextControl.GetPages().GetItem().Header
            RemoveHeaderFooter(hfHeader)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Footer_Insert_Click Handler
        '
        ' Insert a footer to the TextControl (or activates the footer if it already exists).
        '  
        ' Item: 'Insert' of the 'Footer' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Footer_Insert_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            InsertHeaderFooter(HeaderFooterType.Footer)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Footer_Remove_Click Handler
        '
        ' Removes the footer from the TextControl.
        '  
        ' Item: 'Remove' of the 'Footer' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Footer_Remove_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim hfFooter As HeaderFooter = Me.m_txTextControl.GetPages().GetItem().Footer
            RemoveHeaderFooter(hfFooter)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_PageNumber_Insert_Click Handler
        '
        ' Inserts a new page number field at the current input position of the header or footer.
        '  
        ' Item: 'Insert' of the 'Page Number' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_PageNumber_Insert_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim hfHeaderFooter As HeaderFooter = TryCast(Me.m_txTextControl.TextParts.GetItem(), HeaderFooter)
            hfHeaderFooter.PageNumberFields.Add(New PageNumberField())
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_PageNumber_Remove_Click Handler
        '
        ' Removes a page number from the header or footer.
        '  
        ' Item: 'Remove' of the 'Page Number' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_PageNumber_Remove_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim hfHeaderFooter As HeaderFooter = TryCast(Me.m_txTextControl.TextParts.GetItem(), HeaderFooter)
            Dim pnfPageNumberField As PageNumberField = hfHeaderFooter.PageNumberFields.GetItem()
            hfHeaderFooter.PageNumberFields.Remove(pnfPageNumberField)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_FormFields_TextFormField_Click Handler
        '
        ' Inserts a text form field at the current input position.
        '  
        ' Item: 'Text Form Field' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_FormFields_TextFormField_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim tffTextFormField As TextFormField = New TextFormField(GetWidthFromEnvironment())
            Me.m_txTextControl.FormFields.Add(tffTextFormField)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_FormFields_CheckBox_Click Handler
        '
        ' Inserts a check form field at the current input position.
        '  
        ' Item: 'Check Box' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_FormFields_CheckBox_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim cffCheckFormField As CheckFormField = New CheckFormField(True)
            Me.m_txTextControl.FormFields.Add(cffCheckFormField)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_FormFields_ComboBox_Click Handler
        '
        ' Inserts a selection form field (with combo box behavior) at the current input position.
        '  
        ' Item: 'Combo Box' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_FormFields_ComboBox_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim sffComboBoxField As SelectionFormField = New SelectionFormField(GetWidthFromEnvironment()) With {
                .IsDropDownArrowVisible = True
            }
            Me.m_txTextControl.FormFields.Add(sffComboBoxField)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_FormFields_DropDownList_Click Handler
        '
        ' Inserts a selection form field (with drop down list behavior) at the current input position.
        '  
        ' Item: 'Drop-Down List' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_FormFields_DropDownList_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim sffDropDownListField As SelectionFormField = New SelectionFormField(GetWidthFromEnvironment()) With {
                .Editable = False,
                .IsDropDownArrowVisible = True
            }
            Me.m_txTextControl.FormFields.Add(sffDropDownListField)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_FormFields_DateFormField_Click Handler
        '
        ' Inserts a date form field at the current input position.
        '  
        ' Item: 'Date Form Field' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_FormFields_DateFormField_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim dffDateFormField As DateFormField = New DateFormField(0) With {
                .IsDateControlVisible = True
            }
            Me.m_txTextControl.FormFields.Add(dffDateFormField)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_FormFields_Delete_Click Handler
        '
        ' Removes the current form field from the TextControl document.
        '  
        ' Item: 'Delete' of the 'Form Fields' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_FormFields_Delete_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.FormFields.Remove(Me.m_txTextControl.FormFields.GetItem())
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Symbol_Click Handler
        '
        ' Invokes a built-in dialog box for inserting symbol characters.
        '  
        ' Item: 'Symbol...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Symbol_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.AddSymbolDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Hyperlink_Insert_Click Handler
        '
        ' Opens an instance of the HyperlinkDialog to insert a hyperlink (HypertextLink or DocumentLink) at the 
        ' current input position.
        '  
        ' Item: 'Hyperlink...'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Hyperlink_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim dlgHyperlink As HyperlinkDialog = New HyperlinkDialog(Me.m_txTextControl) With {
                .Owner = Me
            }
            dlgHyperlink.ShowDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Bookmark_Insert_Click Handler
        '
        ' Opens an instance of the BookmarkDialog to insert a DocumentTarget at the current input position.
        '  
        ' Item: 'Insert...' of the 'Bookmark' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Bookmark_Insert_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim dlgInsertBookmark As BookmarkDialog = New BookmarkDialog(Me.m_txTextControl) With {
                .Owner = Me
            }
            dlgInsertBookmark.ShowDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Bookmark_Delete_Click Handler
        '
        ' Opens an instance of the DeleteBookmarksDialog to delete document targets from the TextControl document.
        '  
        ' Item: 'Delete...' of the 'Bookmark' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Bookmark_Delete_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim dlgDeleteBookmark As DeleteBookmarksDialog = New DeleteBookmarksDialog(Me.m_txTextControl) With {
                .Owner = Me
            }
            dlgDeleteBookmark.ShowDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_TableOfContents_Insert_Click Handler
        '
        ' Invokes the built-in dialog box to insert a table of contents. 
        '  
        ' Item: 'Insert...' of the 'Table of Contents' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_TableOfContents_Insert_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TableOfContentsDialog()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_TableOfContents_Delete_Click Handler
        '
        ' Removes the table of contents from the collection including all its text and including all DocumentTargets 
        ' to where the table's links point.
        '  
        ' Item: 'Delete' of the 'Table of Contents' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_TableOfContents_Delete_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim tocTableOfContens As TableOfContents = Me.m_txTextControl.TablesOfContents.GetItem()

            If tocTableOfContens IsNot Nothing Then
                Me.m_txTextControl.TablesOfContents.Remove(tocTableOfContens)
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_TableOfContents_Update_Click Handler
        '
        ' Updates the content and the page numbers of the table of contents.
        '  
        ' Item: 'Update' of the 'Table of Contents' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_TableOfContents_Update_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim tocTableOfContens As TableOfContents = Me.m_txTextControl.TablesOfContents.GetItem()

            If tocTableOfContens IsNot Nothing Then
                If tocTableOfContens.Update() = TableOfContentsCollection.AddResult.ContentNotFound Then
                    MessageBox.Show(Me, My.Resources.MessageBox_UpdateTableOfContents_NoContents_Text, My.Resources.MessageBox_UpdateTableOfContents_NoContents_Caption, MessageBoxButton.OK, MessageBoxImage.Information)
                End If
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Columns_One_Click Handler
        '
        ' Set the number of colums for the current section to one.
        '  
        ' Item: 'One' of the 'Columns' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Columns_One_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            SetColumnCount(1)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Columns_Two_Click Handler
        '
        ' Set the number of colums for the current section to two.
        '  
        ' Item: 'Two' of the 'Columns' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Columns_Two_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            SetColumnCount(2)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_Columns_MoreColumns_Click Handler
        '
        ' Opens the third tab of the built-in tabbed dialog box for setting the number of page columns and its 
        ' attributes.
        '  
        ' Item: 'More Columns...' of the 'Columns' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_Columns_MoreColumns_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.SectionFormatDialog(2)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_PageBreaks_Page_Click Handler
        '
        ' Add a page break at the current text position and scrolls to input position on the next page.
        '  
        ' Item: 'Page' of the 'Page Breaks' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_PageBreaks_Page_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TextChars.Add(ControlChars.PageBreak)
            Me.m_txTextControl.ScrollLocation = New Point(Me.m_txTextControl.InputPosition.Location.X, Me.m_txTextControl.InputPosition.Location.Y - 1440)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_PageBreaks_Column_Click Handler
        '
        ' Add a column break at the current text position and scrolls to the new input position.
        '  
        ' Item: 'Column' of the 'Page Breaks' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_PageBreaks_Column_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TextChars.Add(ControlChars.ColumnBreak)
            Me.ScrollToTextPosition(Me.m_txTextControl.InputPosition.TextPosition)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_PageBreaks_TextWrapping_Click Handler
        '
        ' Add a line break at the current text position and scrolls to the new input position.
        '  
        ' Item: 'Text Wrapping' of the 'Page Breaks' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_PageBreaks_TextWrapping_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TextChars.Add(ControlChars.LineBreak)
            Me.ScrollToTextPosition(Me.m_txTextControl.InputPosition.TextPosition)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_SectionBreaks_NextPage_Click Handler
        '
        ' Adds a new section on the next page with a new paragraph.
        '  
        ' Item: 'Next Page' of the 'Section Breaks' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_SectionBreaks_NextPage_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Sections.Add(SectionBreakKind.BeginAtNewPage)
            Me.ScrollToTextPosition(Me.m_txTextControl.InputPosition.TextPosition)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Insert_SectionBreaks_Continuous_Click Handler
        '
        ' Adds a new section on the next line with a new paragraph.
        '  
        ' Item: 'Continuous' of the 'Section Breaks' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Insert_SectionBreaks_Continuous_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Sections.Add(SectionBreakKind.BeginAtNewLine)
            Me.ScrollToTextPosition(Me.m_txTextControl.InputPosition.TextPosition)
        End Sub
    End Class
End Namespace
