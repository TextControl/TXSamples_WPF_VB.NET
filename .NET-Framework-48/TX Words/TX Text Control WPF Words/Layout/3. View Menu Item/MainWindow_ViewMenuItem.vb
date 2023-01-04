'-----------------------------------------------------------------------------------------------------------
' MainWindow_ViewMenuItem.vb File
'
' Description: Provides methods to set the layout of the 'View' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' SetViewItemsTexts Method
        '
        ' Sets the texts of the 'View' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetViewItemsTexts()
            m_miView.Header = My.Resources.Item_View_Text

            ' 'Page Layout'
            m_miView_PageLayout.Header = My.Resources.Item_View_PageLayout_Text

            ' 'Draft'
            m_miView_Draft.Header = My.Resources.Item_View_Draft_Text

            ' 'Button Bar'
            m_miView_ButtonBar.Header = My.Resources.Item_View_ButtonBar_Text

            ' 'Status Bar'
            m_miView_StatusBar.Header = My.Resources.Item_View_StatusBar_Text

            ' 'Horizontal Ruler'
            m_miView_HorizontalRuler.Header = My.Resources.Item_View_HorizontalRuler_Text

            ' 'Vertical Ruler'
            m_miView_VerticalRuler.Header = My.Resources.Item_View_VerticalRuler_Text

            ' 'Table Gridlines'
            m_miView_TableGridlines.Header = My.Resources.Item_View_TableGridlines_Text

            ' 'Bookmark Markers'
            m_miView_BookmarkMarkers.Header = My.Resources.Item_View_BookmarkMarkers_Text

            ' 'Text Frame Marker Lines'
            m_miView_TextFrameMarkerLines.Header = My.Resources.Item_View_TextFrameMarkerLines_Text

            ' 'Drawing Marker Lines'
            m_miView_DrawingMarkerLines.Header = My.Resources.Item_View_DrawingMarkerLines_Text

            ' 'Control Chars'
            m_miView_ControlChars.Header = My.Resources.Item_View_ControlChars_Text

            ' 'Editable Regions'
            m_miView_EditableRegions.Header = My.Resources.Item_View_EditableRegions_Text
            m_miView_EditableRegions_Always.Header = My.Resources.Item_View_EditableRegions_Always_Text
            m_miView_EditableRegions_Current.Header = My.Resources.Item_View_EditableRegions_Current_Text
            m_miView_EditableRegions_Never.Header = My.Resources.Item_View_EditableRegions_Never_Text

            ' 'Tracked Changes'
            m_miView_TrackedChanges.Header = My.Resources.Item_View_TrackedChanges_Text

            ' 'Comments'
            m_miView_Comments.Header = My.Resources.Item_View_Comments_Text
            m_miView_Comments_Always.Header = My.Resources.Item_View_Comments_Always_Text
            m_miView_Comments_Current.Header = My.Resources.Item_View_Comments_Current_Text
            m_miView_Comments_Never.Header = My.Resources.Item_View_Comments_Never_Text

            ' 'Zoom'
            m_miView_Zoom.Header = My.Resources.Item_View_Zoom_Text
            Me.SetItemText(Me.m_miView_Zoom_50, My.Resources.Item_View_Zoom_Factor_Text, "_50")
            Me.SetItemText(Me.m_miView_Zoom_75, My.Resources.Item_View_Zoom_Factor_Text, "_75")
            Me.SetItemText(Me.m_miView_Zoom_100, My.Resources.Item_View_Zoom_Factor_Text, "_100")
            Me.SetItemText(Me.m_miView_Zoom_150, My.Resources.Item_View_Zoom_Factor_Text, "15_0")
            Me.SetItemText(Me.m_miView_Zoom_200, My.Resources.Item_View_Zoom_Factor_Text, "_200")
            Me.SetItemText(Me.m_miView_Zoom_300, My.Resources.Item_View_Zoom_Factor_Text, "_300")
            Me.SetItemText(Me.m_miView_Zoom_400, My.Resources.Item_View_Zoom_Factor_Text, "_400")

            ' 'Right to Left Layout'
            m_miView_RightToLeftLayout.Header = My.Resources.Item_View_RightToLeftLayout_Text
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SetViewItemsImages Method
        '
        ' Sets the images of the 'View' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetViewItemsImages()
            ' 'Page Layout'
            Me.SetItemImage(Me.m_miView_PageLayout, RibbonViewTab.RibbonItem.TXITEM_PrintLayout.ToString())

            ' 'Draft'
            Me.SetItemImage(Me.m_miView_Draft, RibbonViewTab.RibbonItem.TXITEM_Draft.ToString())

            ' 'Table Gridlines'
            Me.SetItemImage(Me.m_miView_TableGridlines, RibbonViewTab.RibbonItem.TXITEM_ShowTableGridlines.ToString())

            ' 'Bookmark Markers'
            Me.SetItemImage(Me.m_miView_BookmarkMarkers, RibbonViewTab.RibbonItem.TXITEM_ShowBookmarkMarkers.ToString())

            ' 'Text Frame Marker Lines'
            Me.SetItemImage(Me.m_miView_TextFrameMarkerLines, RibbonViewTab.RibbonItem.TXITEM_ShowTextFrameMarkersLines.ToString())

            ' 'Drawing Marker Lines'
            Me.SetItemImage(Me.m_miView_DrawingMarkerLines, RibbonViewTab.RibbonItem.TXITEM_ShowDrawingFrameMarkersLines.ToString())

            ' 'Control Chars'
            Me.SetItemImage(Me.m_miView_ControlChars, RibbonViewTab.RibbonItem.TXITEM_ShowControlChars.ToString())

            ' 'Editable Regions'
            Me.SetItemImage(Me.m_miView_EditableRegions, RibbonPermissionsTab.RibbonItem.TXITEM_HighlightEditableRegions.ToString())

            ' 'Tracked Changes'
            Me.SetItemImage(Me.m_miView_TrackedChanges, RibbonProofingTab.RibbonItem.TXITEM_TrackChanges.ToString())

            ' 'Comments'
            Me.SetItemImage(Me.m_miView_Comments, RibbonProofingTab.RibbonItem.TXITEM_CommentsViewMode.ToString())

            ' 'Zoom'
            Me.SetItemImage(Me.m_miView_Zoom, RibbonViewTab.RibbonItem.TXITEM_ZoomFactor.ToString())

            ' 'Right to Left Layout'
            Me.m_miView_RightToLeftLayout.Icon = GetSmallIcon("RightToLeft_Small.svg")
        End Sub
    End Class
End Namespace
