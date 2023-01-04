'-----------------------------------------------------------------------------------------------------------
' MainWindow_FileMenuItem_DropDownOpening.vb File
'
' Description: Provides all SubmenuOpened handlers associated with 'File' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' File_SubmenuOpened Handler
        '
        ' Updates the IsEnabled state of 'File' drop down menu items.
        ' 
        ' Item: 'File'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub File_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' 'Recent Files'
            Me.m_miFile_RecentFiles.IsEnabled = m_colRecentFiles.Count > 0

            ' 'Save'
            Me.m_miFile_Save.IsEnabled = m_bIsDirtyDocument AndAlso Not m_bIsUnknownDocument

            ' 'Print'
            Me.m_miFile_Print.IsEnabled = Me.m_txTextControl.CanPrint
        End Sub
    End Class
End Namespace
