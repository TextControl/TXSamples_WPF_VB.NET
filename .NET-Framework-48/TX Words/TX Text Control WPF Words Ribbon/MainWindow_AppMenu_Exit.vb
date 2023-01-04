'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_Exit.vb File
'
' Description:
'     Manages closing the application.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.ComponentModel

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Exit_Click Handler
        ' Closes the application when clicked.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Exit_Click(ByVal sender As Object, ByVal e As EventArgs)
            Close()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' MainWindow_Closing Handler
        ' Invokes the SaveDirtyDocumentOnExit method to handle dirty documents. If the method returns false, the 
        ' closing of the application will be canceled. If the window closing is not canceled, the recent files
        ' are saved to the My.Settings.Default.RecentFiles property.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As CancelEventArgs)
            If Not (CSharpImpl.Assign(e.Cancel, Not SaveDirtyDocumentOnExit())) Then
                ' Save the recent files to the My.Settings.Default.RecentFiles property
                SaveRecentFiles()
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SaveDirtyDocumentOnExit Method
        ' If the document is dirty (unsaved changes were made), a MessageBox is shown where the user can
        ' decide whether closing the application should be canceled or the changed document should (or should 
        ' not) be saved before closing the program.
        '
        ' Return value:    If closing the application should be canceled, the method returns false. 
        '                  Otherwise true.
        '-----------------------------------------------------------------------------------------------------------
        Private Function SaveDirtyDocumentOnExit() As Boolean
            Dim bKeepGoing = True
            If m_bIsDirtyDocument Then
                Dim strMessageBoxTExt = If(m_bIsUnknownDocument, My.Resources.MessageBox_SaveDirtyDocumentOnExit_Untitled, String.Format(My.Resources.MessageBox_SaveDirtyDocumentOnExit_ToFile, m_strActiveDocumentPath))
                Dim mbrDialogResult = MessageBox.Show(Me, strMessageBoxTExt, My.Resources.MessageBox_SaveDirtyDocumentOnExit_Caption, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)
                Select Case mbrDialogResult
                    Case MessageBoxResult.Yes
                        bKeepGoing = Save(m_strActiveDocumentPath)
                    Case MessageBoxResult.Cancel
                        bKeepGoing = False
                End Select
            End If
            Return bKeepGoing
        End Function
    End Class
End Namespace
