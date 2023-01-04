'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_About.vb File
'
' Description:
'      Handles displaying the 'About' sidebar and manages its content.
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
        ' UpdateAboutSidebar Method
        ' Connects a handler to the about viewer's Loaded event.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateAboutSidebar()
            Dim txAboutViewer As TextControl = TryCast(Me.m_sbSidebarLeft.FindName(Sidebar.AboutItem.TXITEM_AboutViewer.ToString()), TextControl)
            AddHandler txAboutViewer.Loaded, AddressOf AboutViewer_Loaded
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' AboutViewer_Loaded Handler
        ' Loads the About.xml into the about viewer.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub AboutViewer_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim txAboutViewer As TextControl = TryCast(sender, TextControl)
            txAboutViewer.Tables.GridLines = False ' Disable table grid lines.
            txAboutViewer.PageSize = New PageSize(17010, 17010) ' Set a size of 30x30cm
            txAboutViewer.Load(Me.m_strFilesDirectory & "About.xml", StreamType.XMLFormat)
            txAboutViewer.XmlEditMode = XmlEditMode.NoValidate
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' About_Click Handler
        ' Shows or hides the 'About' sidebar
        '-----------------------------------------------------------------------------------------------------------
        Private Sub About_CheckedChanged(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If Me.m_sbSidebarLeft IsNot Nothing Then
                If Me.m_rtbtnAbout.IsChecked Then
                    Me.m_sbSidebarLeft.ContentLayout = Sidebar.SidebarContentLayout.About
                    Me.m_sbSidebarLeft.IsShown = True
                    UpdateAboutSidebar()
                Else
                    Me.m_sbSidebarLeft.IsShown = False
                End If
            End If
        End Sub
    End Class
End Namespace
