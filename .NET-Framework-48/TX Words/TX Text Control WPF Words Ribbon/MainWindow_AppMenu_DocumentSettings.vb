'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_DocumentSettings.vb File
'
' Description:
'     Manages showing/hiding the document settings sidebar.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' DocumentSettings_Click Handler
        ' Shows and hides the Document Settings sidebar when the checked state of the corresponding button
        ' changed.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub DocumentSettings_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim bChecked As Boolean
            If Me.m_rtbtnDocumentSettings.IsChecked Then
                bChecked = True
                ' Set the content layout of the sidebar to DocumentSettings when the button is checked.
                Me.m_sbSidebarLeft.ContentLayout = Sidebar.SidebarContentLayout.DocumentSettings
            End If
            ' Show/hide the sidebar.
            Me.m_sbSidebarLeft.IsShown = bChecked
        End Sub
    End Class
End Namespace
