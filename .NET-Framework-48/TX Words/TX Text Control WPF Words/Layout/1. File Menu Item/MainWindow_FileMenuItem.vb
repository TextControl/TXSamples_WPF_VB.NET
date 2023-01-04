'-----------------------------------------------------------------------------------------------------------
' MainWindow_FileMenuItem.vb File
'
' Description: Provides methods to set the layout of the 'File' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Collections.Specialized
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------

        ' 'Open...' item
        Private ReadOnly m_isPDFImportSettings As PDFImportSettings = PDFImportSettings.GenerateTextFrames Or PDFImportSettings.LoadEmbeddedFiles

        ' 'Recent Files' item
        Private Const m_iMaxRecentFiles As Integer = 10
        Private m_colRecentFiles As StringCollection

        ' 'Sign In...' and '[Current user]' items
        Private m_uiCurrentUser As UserInfo = Nothing ' Info about the current user.
        Private m_strUserName As String = "" ' Environment.UserName


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetFileItemsTexts Method
        '
        ' Sets the texts of the 'File' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetFileItemsTexts()
            ' 'File'
            m_miFile.Header = My.Resources.Item_File_Text
            ' 'New'
            m_miFile_New.Header = My.Resources.Item_File_New_Text
            ' 'Open...'
            m_miFile_Open.Header = My.Resources.Item_File_Open_Text

            ' 'Recent Files'
            m_miFile_RecentFiles.Header = My.Resources.Item_File_RecentFiles_Text

            ' 'Save'
            m_miFile_Save.Header = My.Resources.Item_File_Save_Text

            ' 'Save As...'
            m_miFile_SaveAs.Header = My.Resources.Item_File_SaveAs_Text

            ' 'Page Setup...'
            m_miFile_PageSetup.Header = My.Resources.Item_File_PageSetup_Text

            ' 'Print...'
            m_miFile_Print.Header = My.Resources.Item_File_Print_Text

            ' 'Print Quick'
            m_miFile_PrintQuick.Header = My.Resources.Item_File_PrintQuick_Text

            ' 'Sign In...'
            m_miFile_SignIn.Header = My.Resources.Item_File_SignIn_Text

            ' '[Current User]'
            Me.SetItemText(Me.m_miFile_CurrentUser, m_strUserName)
            m_miFile_CurrentUser_AccountSettings.Header = My.Resources.Item_File_CurrentUser_AccountSettings_Text
            m_miFile_CurrentUser_SignOut.Header = My.Resources.Item_File_CurrentUser_SignOut_Text

            ' 'Exit'
            m_miFile_Exit.Header = My.Resources.Item_File_Exit_Text
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SetFileItemsImages Method
        '
        ' Sets the images of the 'File' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetFileItemsImages()
            ' 'New'
            Me.SetItemImage(Me.m_miFile_New, ResourceProvider.FileMenuItem.TXITEM_New.ToString())

            ' 'Open...'
            Me.SetItemImage(Me.m_miFile_Open, ResourceProvider.FileMenuItem.TXITEM_Open.ToString())

            ' 'Save'
            Me.SetItemImage(Me.m_miFile_Save, ResourceProvider.FileMenuItem.TXITEM_Save.ToString())

            ' 'Save As...'
            Me.SetItemImage(Me.m_miFile_SaveAs, ResourceProvider.FileMenuItem.TXITEM_SaveAs.ToString())

            ' 'Page Setup...'
            Me.SetItemImage(Me.m_miFile_PageSetup, RibbonPageLayoutTab.RibbonItem.TXITEM_PageMargins.ToString())

            ' 'Print...'
            Me.SetItemImage(Me.m_miFile_Print, ResourceProvider.FileMenuItem.TXITEM_Print.ToString())

            ' 'Print Quick'
            Me.SetItemImage(Me.m_miFile_PrintQuick, ResourceProvider.FileMenuItem.TXITEM_PrintQuick.ToString())

            ' 'Sign In...'
            Me.SetItemImage(Me.m_miFile_SignIn, ResourceProvider.FileMenuItem.TXITEM_SignIn.ToString())

            ' '[Current User]'
            Me.SetItemImage(Me.m_miFile_CurrentUser, ResourceProvider.FileMenuItem.TXITEM_CurrentUser.ToString())
            Me.SetItemImage(Me.m_miFile_CurrentUser_AccountSettings, ResourceProvider.FileMenuItem.TXITEM_AccountSettings.ToString())
            Me.SetItemImage(Me.m_miFile_CurrentUser_SignOut, ResourceProvider.FileMenuItem.TXITEM_SignOut.ToString())

            ' 'Exit'
            Me.SetItemImage(Me.m_miFile_Exit, ResourceProvider.FileMenuItem.TXITEM_Exit.ToString())
        End Sub
    End Class
End Namespace
