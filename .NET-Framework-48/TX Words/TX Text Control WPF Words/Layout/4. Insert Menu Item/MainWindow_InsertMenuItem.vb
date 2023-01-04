'-----------------------------------------------------------------------------------------------------------
' MainWindow_InsertMenuItem.vb File
'
' Description: Provides methods to set the layout of the 'Insert' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private Const m_iDefaultEmptyWidth As Integer = 2000  ' The actual default empty width of a form field is 2000 twips when setting the default width integer flag '0'


        '-----------------------------------------------------------------------------------------------------------
        ' Shape Types
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' 'Lines' Item
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_Lines As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.Line, Drawing.ShapeType.BentConnector3, Drawing.ShapeType.CurvedConnector3}

        '-----------------------------------------------------------------------------------------------------------
        ' 'Rectangles' Items
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_Rectangles As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.Rectangle, Drawing.ShapeType.RoundRectangle, Drawing.ShapeType.Snip1Rectangle, Drawing.ShapeType.Snip2SameRectangle, Drawing.ShapeType.Snip2DiagonalRectangle, Drawing.ShapeType.SnipRoundRectangle, Drawing.ShapeType.Round1Rectangle, Drawing.ShapeType.Round2SameRectangle, Drawing.ShapeType.Round2DiagonalRectangle}

        '-----------------------------------------------------------------------------------------------------------
        ' 'Basic Shapes' Items
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_BasicShapes As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.Ellipse, Drawing.ShapeType.Triangle, Drawing.ShapeType.RightTriangle, Drawing.ShapeType.Parallelogram, Drawing.ShapeType.NonIsoscelesTrapezoid, Drawing.ShapeType.Diamond, Drawing.ShapeType.Pentagon, Drawing.ShapeType.Hexagon, Drawing.ShapeType.Heptagon, Drawing.ShapeType.Octagon, Drawing.ShapeType.Decagon, Drawing.ShapeType.Dodecagon, Drawing.ShapeType.Pie, Drawing.ShapeType.Chord, Drawing.ShapeType.Teardrop, Drawing.ShapeType.Frame, Drawing.ShapeType.HalfFrame, Drawing.ShapeType.Corner, Drawing.ShapeType.DiagonalStripe, Drawing.ShapeType.Plus, Drawing.ShapeType.Plaque, Drawing.ShapeType.Can, Drawing.ShapeType.Cube, Drawing.ShapeType.Bevel, Drawing.ShapeType.Donut, Drawing.ShapeType.NoSmoking, Drawing.ShapeType.BlockArc, Drawing.ShapeType.FoldedCorner, Drawing.ShapeType.SmileyFace, Drawing.ShapeType.Heart, Drawing.ShapeType.LightningBolt, Drawing.ShapeType.Sun, Drawing.ShapeType.Moon, Drawing.ShapeType.Cloud, Drawing.ShapeType.Arc, Drawing.ShapeType.BracketPair, Drawing.ShapeType.BracePair, Drawing.ShapeType.LeftBracket, Drawing.ShapeType.RightBracket, Drawing.ShapeType.LeftBrace, Drawing.ShapeType.RightBrace}

        '-----------------------------------------------------------------------------------------------------------
        ' 'Block Arrows' Items
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_BlockArrows As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.RightArrow, Drawing.ShapeType.LeftArrow, Drawing.ShapeType.UpArrow, Drawing.ShapeType.DownArrow, Drawing.ShapeType.LeftRightArrow, Drawing.ShapeType.UpDownArrow, Drawing.ShapeType.QuadArrow, Drawing.ShapeType.LeftRightUpArrow, Drawing.ShapeType.BentArrow, Drawing.ShapeType.UTurnArrow, Drawing.ShapeType.LeftUpArrow, Drawing.ShapeType.BentUpArrow, Drawing.ShapeType.CurvedRightArrow, Drawing.ShapeType.CurvedLeftArrow, Drawing.ShapeType.CurvedUpArrow, Drawing.ShapeType.CurvedDownArrow, Drawing.ShapeType.StripedRightArrow, Drawing.ShapeType.NotchedRightArrow, Drawing.ShapeType.NotchedRightArrow, Drawing.ShapeType.Chevron, Drawing.ShapeType.RightArrowCallout, Drawing.ShapeType.DownArrowCallout, Drawing.ShapeType.LeftArrowCallout, Drawing.ShapeType.UpArrowCallout, Drawing.ShapeType.LeftRightArrowCallout, Drawing.ShapeType.QuadArrowCallout, Drawing.ShapeType.CircularArrow}


        '-----------------------------------------------------------------------------------------------------------
        ' 'Equation Shapes' Items
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_EquationShapes As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.MathPlus, Drawing.ShapeType.MathMinus, Drawing.ShapeType.MathMultiply, Drawing.ShapeType.MathDivide, Drawing.ShapeType.MathEqual, Drawing.ShapeType.MathNotEqual}

        '-----------------------------------------------------------------------------------------------------------
        ' 'Flowchart' Items
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_Flowchart As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.FlowChartProcess, Drawing.ShapeType.FlowChartAlternateProcess, Drawing.ShapeType.FlowChartDecision, Drawing.ShapeType.FlowChartInputOutput, Drawing.ShapeType.FlowChartPredefinedProcess, Drawing.ShapeType.FlowChartInternalStorage, Drawing.ShapeType.FlowChartDocument, Drawing.ShapeType.FlowChartMultidocument, Drawing.ShapeType.FlowChartTerminator, Drawing.ShapeType.FlowChartPreparation, Drawing.ShapeType.FlowChartManualInput, Drawing.ShapeType.FlowChartManualOperation, Drawing.ShapeType.FlowChartConnector, Drawing.ShapeType.FlowChartOffpageConnector, Drawing.ShapeType.FlowChartPunchedCard, Drawing.ShapeType.FlowChartPunchedTape, Drawing.ShapeType.FlowChartSummingJunction, Drawing.ShapeType.FlowChartOr, Drawing.ShapeType.FlowChartCollate, Drawing.ShapeType.FlowChartSort, Drawing.ShapeType.FlowChartExtract, Drawing.ShapeType.FlowChartMerge, Drawing.ShapeType.FlowChartOnlineStorage, Drawing.ShapeType.FlowChartDelay, Drawing.ShapeType.FlowChartMagneticTape, Drawing.ShapeType.FlowChartMagneticDisk, Drawing.ShapeType.FlowChartMagneticDrum, Drawing.ShapeType.FlowChartDisplay}

        '-----------------------------------------------------------------------------------------------------------
        ' 'Stars and Banners' Items
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_StarsAndBanners As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.IrregularSeal1, Drawing.ShapeType.IrregularSeal2, Drawing.ShapeType.Star4, Drawing.ShapeType.Star5, Drawing.ShapeType.Star6, Drawing.ShapeType.Star7, Drawing.ShapeType.Star8, Drawing.ShapeType.Star10, Drawing.ShapeType.Star12, Drawing.ShapeType.Star16, Drawing.ShapeType.Star24, Drawing.ShapeType.Star32, Drawing.ShapeType.Ribbon2, Drawing.ShapeType.Ribbon, Drawing.ShapeType.EllipseRibbon2, Drawing.ShapeType.EllipseRibbon, Drawing.ShapeType.VerticalScroll, Drawing.ShapeType.HorizontalScroll, Drawing.ShapeType.Wave, Drawing.ShapeType.DoubleWave}

        '-----------------------------------------------------------------------------------------------------------
        ' 'Callouts' Items
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rstShape_Callouts As Drawing.ShapeType() = New Drawing.ShapeType() {Drawing.ShapeType.WedgeRectangleCallout, Drawing.ShapeType.WedgeRoundRectangleCallout, Drawing.ShapeType.WedgeEllipseCallout, Drawing.ShapeType.CloudCallout, Drawing.ShapeType.BorderCallout1, Drawing.ShapeType.BorderCallout2, Drawing.ShapeType.BorderCallout3, Drawing.ShapeType.AccentCallout1, Drawing.ShapeType.AccentCallout2, Drawing.ShapeType.AccentCallout3, Drawing.ShapeType.Callout1, Drawing.ShapeType.Callout2, Drawing.ShapeType.Callout3, Drawing.ShapeType.AccentBorderCallout1, Drawing.ShapeType.AccentBorderCallout2, Drawing.ShapeType.AccentBorderCallout3}


        '-----------------------------------------------------------------------------------------------------------
        ' Barcode Types
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rbtBarcodeTypes As Barcode.BarcodeType() = New Barcode.BarcodeType() {Barcode.BarcodeType.QRCode, Barcode.BarcodeType.Code128, Barcode.BarcodeType.EAN13, Barcode.BarcodeType.UPCA, Barcode.BarcodeType.EAN8, Barcode.BarcodeType.Interleaved2of5, Barcode.BarcodeType.Postnet, Barcode.BarcodeType.Code39, Barcode.BarcodeType.AztecCode, Barcode.BarcodeType.IntelligentMail, Barcode.BarcodeType.Datamatrix, Barcode.BarcodeType.PDF417, Barcode.BarcodeType.MicroPDF, Barcode.BarcodeType.Codabar, Barcode.BarcodeType.FourState, Barcode.BarcodeType.Code11, Barcode.BarcodeType.Code93, Barcode.BarcodeType.PLANET, Barcode.BarcodeType.RoyalMail, Barcode.BarcodeType.Maxicode}



        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetInsertItemsTexts Method
        '
        ' Sets the texts of the 'Insert' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetInsertItemsTexts()
            ' 'Insert'
            m_miInsert.Header = My.Resources.Item_Insert_Text

            ' 'File...'
            m_miInsert_File.Header = My.Resources.Item_Insert_File_Text

            ' 'Image...'
            m_miInsert_Image.Header = My.Resources.Item_Insert_Image_Text

            ' 'Text Frame'
            m_miInsert_TextFrame.Header = My.Resources.Item_Insert_TextFrame_Text

            ' 'Shape'	
            m_miInsert_Shape.Header = My.Resources.Item_Insert_Shape_Text
            m_miInsert_Shape_Lines.Header = My.Resources.Item_Insert_Shape_Lines_Text
            Me.SetItemText(Me.m_miInsert_Shape_Lines)
            m_miInsert_Shape_Rectangles.Header = My.Resources.Item_Insert_Shape_Rectangles_Text
            Me.SetItemText(Me.m_miInsert_Shape_Rectangles)
            m_miInsert_Shape_BasicShapes.Header = My.Resources.Item_Insert_Shape_BasicShapes_Text
            Me.SetItemText(Me.m_miInsert_Shape_BasicShapes)
            m_miInsert_Shape_BlockArrows.Header = My.Resources.Item_Insert_Shape_BlockArrows_Text
            Me.SetItemText(Me.m_miInsert_Shape_BlockArrows)
            m_miInsert_Shape_EquationShapes.Header = My.Resources.Item_Insert_Shape_EquationShapes_Text
            Me.SetItemText(Me.m_miInsert_Shape_EquationShapes)
            m_miInsert_Shape_Flowchart.Header = My.Resources.Item_Insert_Shape_Flowchart_Text
            Me.SetItemText(Me.m_miInsert_Shape_Flowchart)
            m_miInsert_Shape_StarsAndBanners.Header = My.Resources.Item_Insert_Shape_StarsAndBanners_Text
            Me.SetItemText(Me.m_miInsert_Shape_StarsAndBanners)
            m_miInsert_Shape_Callouts.Header = My.Resources.Item_Insert_Shape_Callouts_Text
            Me.SetItemText(Me.m_miInsert_Shape_Callouts)
            Me.SetItemText(Me.m_miInsert_Shape_DrawingCanvas)
            m_miInsert_Shape_DrawingCanvas.Header = My.Resources.Item_Insert_Shape_DrawingCanvas_Text

            ' 'Barcode'
            m_miInsert_Barcode.Header = My.Resources.Item_Insert_Barcode_Text
            Me.SetItemText(Me.m_miInsert_Barcode)

            ' 'Header'
            m_miInsert_Header.Header = My.Resources.Item_Insert_Header_Text
            m_miInsert_Header_Insert.Header = My.Resources.Item_Insert_Header_Insert_Text
            m_miInsert_Header_Remove.Header = My.Resources.Item_Insert_Header_Remove_Text

            ' 'Footer'
            m_miInsert_Footer_Insert.Header = My.Resources.Item_Insert_Footer_Text
            m_miInsert_Footer_Remove.Header = My.Resources.Item_Insert_Footer_Insert_Text

            ' 'Page Number'
            m_miInsert_PageNumber.Header = My.Resources.Item_Insert_PageNumber_Text
            m_miInsert_PageNumber_Insert.Header = My.Resources.Item_Insert_PageNumber_Insert_Text
            m_miInsert_PageNumber_Delete.Header = My.Resources.Item_Insert_PageNumber_Delete_Text

            ' 'Form Fields'
            m_miInsert_FormField.Header = My.Resources.Item_Insert_FormField_Text
            m_miInsert_FormField_TextFormField.Header = My.Resources.Item_Insert_FormField_TextFormField_Text
            m_miInsert_FormField_CheckBox.Header = My.Resources.Item_Insert_FormField_CheckBox_Text
            m_miInsert_FormField_ComboBox.Header = My.Resources.Item_Insert_FormField_ComboBox_Text
            m_miInsert_FormField_DropDownList.Header = My.Resources.Item_Insert_FormField_DropDownList_Text
            m_miInsert_FormField_DateFormField.Header = My.Resources.Item_Insert_FormField_DateFormField_Text
            m_miInsert_FormField_Delete.Header = My.Resources.Item_Insert_FormField_Delete_Text

            ' 'Symbol'
            m_miInsert_Symbol.Header = My.Resources.Item_Insert_Symbol_Text

            ' 'Hyperlink...'
            m_miInsert_Hyperlink.Header = My.Resources.Item_Insert_Hyperlink_Text

            ' 'Bookmark...'
            m_miInsert_Bookmark.Header = My.Resources.Item_Insert_Bookmark_Text
            m_miInsert_Bookmark_Insert.Header = My.Resources.Item_Insert_Bookmark_Insert_Text
            m_miInsert_Bookmark_Delete.Header = My.Resources.Item_Insert_Bookmark_Delete_Text

            ' 'Table of Contents'
            m_miInsert_TableOfContents.Header = My.Resources.Item_Insert_TableOfContents_Text
            m_miInsert_TableOfContents_Insert.Header = My.Resources.Item_Insert_TableOfContents_Insert_Text
            m_miInsert_TableOfContents_Delete.Header = My.Resources.Item_Insert_TableOfContents_Delete_Text
            m_miInsert_TableOfContents_Update.Header = My.Resources.Item_Insert_TableOfContents_Update_Text

            ' 'Columns'
            m_miInsert_Columns.Header = My.Resources.Item_Insert_Columns_Text
            m_miInsert_Columns_One.Header = My.Resources.Item_Insert_Columns_One_Text
            m_miInsert_Columns_Two.Header = My.Resources.Item_Insert_Columns_Two_Text
            m_miInsert_Columns_MoreColumns.Header = My.Resources.Item_Insert_Columns_MoreColumns_Text

            ' 'Page Breaks'
            m_miInsert_PageBreaks.Header = My.Resources.Item_Insert_PageBreaks_Text
            m_miInsert_PageBreaks_Page.Header = My.Resources.Item_Insert_PageBreaks_Page_Text
            m_miInsert_PageBreaks_Column.Header = My.Resources.Item_Insert_PageBreaks_Column_Text
            m_miInsert_PageBreaks_TextWrapping.Header = My.Resources.Item_Insert_PageBreaks_TextWrapping_Text

            ' 'Section Breaks'
            m_miInsert_SectionBreaks.Header = My.Resources.Item_Insert_SectionBreaks_Text
            m_miInsert_SectionBreaks_NextPage.Header = My.Resources.Item_Insert_SectionBreaks_NextPage_Text
            m_miInsert_SectionBreaks_Continuous.Header = My.Resources.Item_Insert_SectionBreaks_Continuous_Text
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SetInsertItemsImages Method
        '
        ' Sets the images of the 'Insert' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetInsertItemsImages()
            ' 'File...'
            Me.SetItemImage(Me.m_miInsert_File, RibbonInsertTab.RibbonItem.TXITEM_InsertFile.ToString())

            ' 'Image...'
            Me.SetItemImage(Me.m_miInsert_Image, RibbonInsertTab.RibbonItem.TXITEM_InsertImage.ToString())

            ' 'Text Frame'
            Me.SetItemImage(Me.m_miInsert_TextFrame, RibbonInsertTab.RibbonItem.TXITEM_InsertTextFrame.ToString())

            ' 'Shape'	
            Me.SetItemImage(Me.m_miInsert_Shape, RibbonInsertTab.RibbonItem.TXITEM_InsertShape.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_Lines, ResourceProvider.ShapeItem.TXITEM_SHAPE_Line.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_Rectangles, ResourceProvider.ShapeItem.TXITEM_SHAPE_Rectangle.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_BasicShapes, ResourceProvider.ShapeItem.TXITEM_SHAPE_RightTriangle.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_BlockArrows, ResourceProvider.ShapeItem.TXITEM_SHAPE_RightArrow.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_EquationShapes, ResourceProvider.ShapeItem.TXITEM_SHAPE_Plus.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_Flowchart, ResourceProvider.ShapeItem.TXITEM_SHAPE_FlowChartMultidocument.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_StarsAndBanners, ResourceProvider.ShapeItem.TXITEM_SHAPE_Star7.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_Callouts, ResourceProvider.ShapeItem.TXITEM_SHAPE_WedgeRoundRectangleCallout.ToString())
            Me.SetItemImage(Me.m_miInsert_Shape_DrawingCanvas, RibbonInsertTab.RibbonDropDownItem.TXITEM_InsertDrawingCanvas.ToString())

            ' 'Barcode'	
            Me.SetItemImage(Me.m_miInsert_Barcode, RibbonInsertTab.RibbonItem.TXITEM_InsertBarcode.ToString())

            ' 'Header'
            Me.SetItemImage(Me.m_miInsert_Header, RibbonInsertTab.RibbonItem.TXITEM_InsertHeader.ToString())
            Me.SetItemImage(Me.m_miInsert_Header_Insert, RibbonInsertTab.RibbonDropDownItem.TXITEM_EditHeader.ToString())
            Me.SetItemImage(Me.m_miInsert_Header_Remove, RibbonInsertTab.RibbonDropDownItem.TXITEM_RemoveHeader.ToString())

            ' 'Footer'
            Me.SetItemImage(Me.m_miInsert_Footer, RibbonInsertTab.RibbonItem.TXITEM_InsertFooter.ToString())
            Me.SetItemImage(Me.m_miInsert_Footer_Insert, RibbonInsertTab.RibbonDropDownItem.TXITEM_EditFooter.ToString())
            Me.SetItemImage(Me.m_miInsert_Footer_Remove, RibbonInsertTab.RibbonDropDownItem.TXITEM_RemoveFooter.ToString())

            ' 'Page Number'
            Me.SetItemImage(Me.m_miInsert_PageNumber, RibbonInsertTab.RibbonItem.TXITEM_InsertPage.ToString())
            Me.SetItemImage(Me.m_miInsert_PageNumber_Insert, RibbonInsertTab.RibbonDropDownItem.TXITEM_InsertStandardPageNumber.ToString())
            Me.SetItemImage(Me.m_miInsert_PageNumber_Delete, RibbonInsertTab.RibbonDropDownItem.TXITEM_RemovePageNumber.ToString())

            ' 'Form Fields'
            Me.SetItemImage(Me.m_miInsert_FormField, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertComboBoxField.ToString())
            Me.SetItemImage(Me.m_miInsert_FormField_TextFormField, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertTextFormField.ToString())
            Me.SetItemImage(Me.m_miInsert_FormField_CheckBox, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertCheckBoxField.ToString())
            Me.SetItemImage(Me.m_miInsert_FormField_ComboBox, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertComboBoxField.ToString())
            Me.SetItemImage(Me.m_miInsert_FormField_DropDownList, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertDropDownListField.ToString())
            Me.SetItemImage(Me.m_miInsert_FormField_DateFormField, RibbonFormFieldsTab.RibbonItem.TXITEM_InsertDateFormField.ToString())
            Me.SetItemImage(Me.m_miInsert_FormField_Delete, RibbonFormFieldsTab.RibbonItem.TXITEM_DeleteFormField.ToString())

            ' 'Symbol'
            Me.SetItemImage(Me.m_miInsert_Symbol, RibbonInsertTab.RibbonItem.TXITEM_InsertSymbol.ToString())

            ' 'Hyperlink...'
            Me.SetItemImage(Me.m_miInsert_Hyperlink, RibbonInsertTab.RibbonItem.TXITEM_InsertHyperlink.ToString())

            ' 'Bookmark...'
            Me.SetItemImage(Me.m_miInsert_Bookmark, RibbonInsertTab.RibbonItem.TXITEM_InsertBookmark.ToString())
            Me.SetItemImage(Me.m_miInsert_Bookmark_Insert, RibbonInsertTab.RibbonDropDownItem.TXITEM_EditBookmark.ToString())
            Me.SetItemImage(Me.m_miInsert_Bookmark_Delete, RibbonInsertTab.RibbonDropDownItem.TXITEM_DeleteBookmark.ToString())

            ' 'Table of Contents'
            Me.SetItemImage(Me.m_miInsert_TableOfContents, RibbonReferencesTab.RibbonItem.TXITEM_InsertTableOfContents.ToString())
            Me.SetItemImage(Me.m_miInsert_TableOfContents_Insert, RibbonReferencesTab.RibbonItem.TXITEM_InsertTableOfContents.ToString())
            Me.SetItemImage(Me.m_miInsert_TableOfContents_Delete, RibbonReferencesTab.RibbonItem.TXITEM_DeleteTableOfContents.ToString())
            Me.SetItemImage(Me.m_miInsert_TableOfContents_Update, RibbonReferencesTab.RibbonItem.TXITEM_UpdateTableOfContents.ToString())

            ' 'Columns'
            Me.SetItemImage(Me.m_miInsert_Columns, RibbonPageLayoutTab.RibbonItem.TXITEM_Columns.ToString())
            Me.SetItemImage(Me.m_miInsert_Columns_One, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Columns_One.ToString())
            Me.SetItemImage(Me.m_miInsert_Columns_Two, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Columns_Two.ToString())
            Me.SetItemImage(Me.m_miInsert_Columns_MoreColumns, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Columns_MoreColumns.ToString())

            ' 'Page Breaks'
            Me.SetItemImage(Me.m_miInsert_PageBreaks, RibbonPageLayoutTab.RibbonItem.TXITEM_Breaks.ToString())
            Me.SetItemImage(Me.m_miInsert_PageBreaks_Page, RibbonPageLayoutTab.RibbonItem.TXITEM_Breaks.ToString())
            Me.SetItemImage(Me.m_miInsert_PageBreaks_Column, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Breaks_Column.ToString())
            Me.SetItemImage(Me.m_miInsert_PageBreaks_TextWrapping, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Breaks_TextWrapping.ToString())

            ' 'Section Breaks'
            Me.SetItemImage(Me.m_miInsert_SectionBreaks, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Breaks_NextPage.ToString())
            Me.SetItemImage(Me.m_miInsert_SectionBreaks_NextPage, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Breaks_NextPage.ToString())
            Me.SetItemImage(Me.m_miInsert_SectionBreaks_Continuous, RibbonPageLayoutTab.RibbonDropDownItem.TXITEM_Breaks_Continuous.ToString())
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' CreateShapeAndBarcodeItems Method
        '
        ' Creates Shape and Barcode items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub CreateShapeAndBarcodeItems()
            ' 'Shape'	
            For Each shapeType In m_rstShape_Lines             ' 'Lines'
                Me.AddShapeItem(Me.m_miInsert_Shape_Lines.Items, shapeType)
            Next

            For Each shapeType In m_rstShape_Rectangles        ' 'Rectangles'
                Me.AddShapeItem(Me.m_miInsert_Shape_Rectangles.Items, shapeType)
            Next

            For Each shapeType In m_rstShape_BasicShapes       ' 'Basic Shapes'
                Me.AddShapeItem(Me.m_miInsert_Shape_BasicShapes.Items, shapeType)
            Next

            For Each shapeType In m_rstShape_BlockArrows       ' 'Block Arrows'
                Me.AddShapeItem(Me.m_miInsert_Shape_BlockArrows.Items, shapeType)
            Next

            For Each shapeType In m_rstShape_EquationShapes    ' 'Equation Shapes'
                Me.AddShapeItem(Me.m_miInsert_Shape_EquationShapes.Items, shapeType)
            Next

            For Each shapeType In m_rstShape_Flowchart         ' 'Flowchart'
                Me.AddShapeItem(Me.m_miInsert_Shape_Flowchart.Items, shapeType)
            Next

            For Each shapeType In m_rstShape_StarsAndBanners   ' 'Stars and Banners'
                Me.AddShapeItem(Me.m_miInsert_Shape_StarsAndBanners.Items, shapeType)
            Next

            For Each shapeType In m_rstShape_Callouts          ' 'Callouts'
                Me.AddShapeItem(Me.m_miInsert_Shape_Callouts.Items, shapeType)
            Next

            ' 'Barcode'	
            For Each barcodeType In m_rbtBarcodeTypes
                Me.AddBarcodeItem(Me.m_miInsert_Barcode.Items, barcodeType)
            Next
        End Sub
    End Class
End Namespace
