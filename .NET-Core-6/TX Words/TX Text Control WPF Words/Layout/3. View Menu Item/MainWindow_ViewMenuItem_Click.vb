'-----------------------------------------------------------------------------------------------------------
' MainWindow_ViewMenuItem_Click.vb File
'
' Description: Provides all Click handlers associated with 'View' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' View_PageLayout_Click Handler
        '
        ' Enables a view mode where the text is formatted according to the settings of the PageSize and the 
        ' PageMargins properties. Additionally the TextControl displays the pages in 3D view with gaps and a desktop
        ' background.
        ' 
        ' Item: 'Page Layout'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_PageLayout_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.ViewMode = ViewMode.PageView
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_Draft_Click Handler
        ' 
        ' Enables a view mode where the text is formatted according to the settings of the PageSize and the 
        ' PageMargins properties.
        '
        ' Item: 'Draft'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_Draft_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.ViewMode = ViewMode.Normal
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_ButtonBar_Click Handler
        ' 
        ' Displays or hides the ButtonBar control that can be used to show or to set font and paragraph attributes   
        ' of theTextControl. 
        '
        ' Item: 'Button Bar'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_ButtonBar_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_bbButtonBar.Visibility = If(Me.m_miView_ButtonBar.IsChecked, Visibility.Visible, Visibility.Collapsed)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_StatusBar_Click Handler
        ' 
        ' Displays or hides the StatusBar control that can be used to show the position of the current text input 
        ' position and other status information of the TextControl. 
        '
        ' Item: 'Status Bar'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_StatusBar_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_sbStatusBar.Visibility = If(Me.m_miView_StatusBar.IsChecked, Visibility.Visible, Visibility.Collapsed)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_HorizontalRuler_Click Handler
        ' 
        ' Displays or hides the horizontal RulerBar control that is connected to the TextControl. 
        '
        ' Item: 'Horizontal Ruler'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_HorizontalRuler_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_rbHorizontalRulerBar.Visibility = If(Me.m_miView_HorizontalRuler.IsChecked, Visibility.Visible, Visibility.Collapsed)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_VerticalRuler_Click Handler
        ' 
        ' Displays or hides the vertical RulerBar control that is connected to the TextControl. 
        '
        ' Item: 'Vertical Ruler'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_VerticalRuler_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_rbVerticalRulerBar.Visibility = If(Me.m_miView_VerticalRuler.IsChecked, Visibility.Visible, Visibility.Collapsed)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_TableGridlines_Click Handler
        ' 
        ' Sets a value whether table gridlines are shown or not.
        '
        ' Item: 'Table Gridlines'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_TableGridlines_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Tables.GridLines = Me.m_miView_TableGridlines.IsChecked
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_BookmarkMarkers_Click Handler
        ' 
        ' Sets a value whether bookmark markers are shown or not.
        '
        ' Item: 'Bookmark Markers'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_BookmarkMarkers_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.DocumentTargetMarkers = Me.m_miView_BookmarkMarkers.IsChecked
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_TextFrameMarkerLines_Click Handler
        ' 
        ' Sets a value whether text frames that have no border line are shown with marker lines.
        '
        ' Item: 'Text Frame Marker Lines'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_TextFrameMarkerLines_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.TextFrameMarkerLines = Me.m_miView_TextFrameMarkerLines.IsChecked
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_DrawingMarkerLines_Click Handler
        ' 
        ' Sets a value whether a marker frame is shown around a drawing to indicate its position and size.
        '
        ' Item: 'Drawing Marker Lines'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_DrawingMarkerLines_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.DrawingMarkerLines = Me.m_miView_DrawingMarkerLines.IsChecked
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_ControlChars_Click Handler
        ' 
        ' Sets a value whether control characters are visible or not.
        '
        ' Item: 'Control Chars'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_ControlChars_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.ControlChars = Me.m_miView_ControlChars.IsChecked
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_EditableRegions_Always_Click Handler
        ' 
        ' Determines that an editable region is always highlighted.
        '
        ' Item: 'Always' of the 'Editable Regions' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_EditableRegions_Always_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.EditableRegionHighlightMode = HighlightMode.Always
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_EditableRegions_Current_Click Handler
        ' 
        ' Determines that an editable region is only highlighted, when it contains the current text input position 
        ' and when the control has the input focus.
        '
        ' Item: 'Current' of the 'Editable Regions' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_EditableRegions_Current_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.EditableRegionHighlightMode = HighlightMode.Activated
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_EditableRegions_Never_Click Handler
        ' 
        ' Determines that an editable region is never highlighted.
        '
        ' Item: 'Never' of the 'Editable Regions' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_EditableRegions_Never_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.EditableRegionHighlightMode = HighlightMode.Never
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_TrackedChanges_Click Handler
        ' 
        ' Sets a value whether a tracked change is always or never highlighted.
        '
        ' Item: 'Tracked Changes'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_TrackedChanges_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.SetTrackedChangesHighlightMode(If(Me.m_miView_TrackedChanges.IsChecked, HighlightMode.Always, HighlightMode.Never))
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_Comments_Always_Click Handler
        ' 
        ' Determines that commented texts are always highlighted.
        '
        ' Item: 'Always' of the 'Comments' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_Comments_Always_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.CommentHighlightMode = HighlightMode.Always
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_Comments_Current_Click Handler
        ' 
        ' Determines that a commented text is only highlighted when it contains the text input position.
        '
        ' Item: 'Current' of the 'Comments' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_Comments_Current_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.CommentHighlightMode = HighlightMode.Activated
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_Comments_Never_Click Handler
        ' 
        ' Determines that no commented text is highlighted.
        '
        ' Item: 'Never' of the 'Comments' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_Comments_Never_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.CommentHighlightMode = HighlightMode.Never
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_Zoom_MenuItem_Click Handler
        ' 
        ' Sets the zoom factor, in percent, for a TextControl. The value is represented by the clicked item's
        ' Tag property value.
        '
        ' Item: Each item of the 'Zoom' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_Zoom_MenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim strZoomItemValue As String = TryCast(e.Source, MenuItem).Tag.ToString()
            Me.m_txTextControl.ZoomFactor = Integer.Parse(strZoomItemValue)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' View_RightToLeftLayout_Click Handler
        ' 
        ' Restarts the application with a program's view that has a reversed text appearance. Furthermore
        ' the user can save the current document before closing the application if the document is dirty.
        '
        ' Item: 'Right to Left Layout'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub View_RightToLeftLayout_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.ReverseAppTextAppearance(Me.m_miView_RightToLeftLayout.IsChecked)
        End Sub
    End Class
End Namespace
