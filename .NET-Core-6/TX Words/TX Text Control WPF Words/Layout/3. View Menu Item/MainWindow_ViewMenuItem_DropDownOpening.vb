'-----------------------------------------------------------------------------------------------------------
' MainWindow_ViewMenuItem_DropDownOpening.vb File
'
' Description: Provides all SubmenuOpened handlers associated with 'View' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' View_SubmenuOpened Handler
        '
        ' Updates the checked state of 'View' drop down menu items.
        ' 
        ' Item: 'View'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' 'Page Layout'
            Me.m_miView_PageLayout.IsChecked = Me.m_txTextControl.ViewMode = ViewMode.PageView

            ' 'Draft'
            Me.m_miView_Draft.IsChecked = Me.m_txTextControl.ViewMode = ViewMode.Normal

            ' 'Button Bar'
            Me.m_miView_ButtonBar.IsChecked = Me.m_bbButtonBar.Visibility = Visibility.Visible

            ' 'Status Bar'
            Me.m_miView_StatusBar.IsChecked = Me.m_sbStatusBar.Visibility = Visibility.Visible

            ' 'Horizontal Ruler'
            Me.m_miView_HorizontalRuler.IsChecked = Me.m_rbHorizontalRulerBar.Visibility = Visibility.Visible

            ' 'Vertical Ruler'
            Me.m_miView_VerticalRuler.IsChecked = Me.m_rbVerticalRulerBar.Visibility = Visibility.Visible

            ' 'Table Gridlines'
            Me.m_miView_TableGridlines.IsChecked = Me.m_txTextControl.Tables.GridLines

            ' 'Bookmark Markers'
            Me.m_miView_BookmarkMarkers.IsChecked = Me.m_txTextControl.DocumentTargetMarkers

            ' 'Text Frame Marker Lines'
            Me.m_miView_TextFrameMarkerLines.IsChecked = Me.m_txTextControl.TextFrameMarkerLines

            ' 'Drawing Marker Lines'
            Me.m_miView_DrawingMarkerLines.IsChecked = Me.m_txTextControl.DrawingMarkerLines

            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' 'Tracked Changes'

                ' Step through all tracked changes to get their common highlight mode.
                Dim colTrackedChanges As TrackedChangeCollection = Me.m_txTextControl.TrackedChanges
                Dim iCount = colTrackedChanges.Count
                Dim hmCurrentHighlightMode = HighlightMode.Always

                For i = 1 To iCount - 1
                    hmCurrentHighlightMode = colTrackedChanges(i).HighlightMode
                    ' Check whether the current tracked change highlight mode differs to the next one's
                    If hmCurrentHighlightMode <> colTrackedChanges(i + 1).HighlightMode Then
                        ' In that case set the 'Tracked Changes' item's checked value to false.
                        Me.m_miView_TrackedChanges.IsChecked = False
                        Return
                    End If
                Next

                ' The 'Tracked Changes' item is checked if the highlight mode of all tracked changes
                ' is set to HighlightMode.Always
                Me.m_miView_TrackedChanges.IsChecked = hmCurrentHighlightMode = HighlightMode.Always
            End If

            ' 'Right to Left Layout'
            Me.m_miView_RightToLeftLayout.IsChecked = FlowDirection = FlowDirection.RightToLeft
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_EditableRegions_SubmenuOpened Handler
        '
        ' Updates the checked state of 'Editable Regions' drop down menu items.
        ' 
        ' Item: 'Editable Regions'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_EditableRegions_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Set the check states of the 'Editable Regions' drop down items.
            Dim hmHighlightMode As HighlightMode = Me.m_txTextControl.EditableRegionHighlightMode
            Me.m_miView_EditableRegions_Always.IsChecked = hmHighlightMode = HighlightMode.Always
            Me.m_miView_EditableRegions_Current.IsChecked = hmHighlightMode = HighlightMode.Activated
            Me.m_miView_EditableRegions_Never.IsChecked = hmHighlightMode = HighlightMode.Never
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_Comments_SubmenuOpened Handler
        '
        ' Updates the checked state of 'Comments' drop down menu items.
        ' 
        ' Item: 'Comments'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_Comments_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Set the check states of the 'Comments' drop down items.
            Dim hmHighlightMode As HighlightMode = Me.m_txTextControl.CommentHighlightMode
            Me.m_miView_Comments_Always.IsChecked = hmHighlightMode = HighlightMode.Always
            Me.m_miView_Comments_Current.IsChecked = hmHighlightMode = HighlightMode.Activated
            Me.m_miView_Comments_Never.IsChecked = hmHighlightMode = HighlightMode.Never
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_Zoom_SubmenuOpened Handler
        '
        ' Updates the checked state of 'Zoom' drop down menu items.
        ' 
        ' Item: 'Zoom'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_Zoom_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Set the check states of the 'Zoom' drop down items.
            Dim strZoomFactor As String = Me.m_txTextControl.ZoomFactor.ToString()

            For Each item As MenuItem In Me.m_miView_Zoom.Items
                item.IsChecked = Equals(item.Tag.ToString(), strZoomFactor)
            Next
        End Sub
    End Class
End Namespace
