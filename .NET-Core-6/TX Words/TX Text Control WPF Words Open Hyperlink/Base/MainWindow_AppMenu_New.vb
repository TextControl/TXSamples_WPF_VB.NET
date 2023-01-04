'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_New.vb File
'
' Description:
'     Manages resetting the content of the document.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' New_Click Handler
        ' Invokes the TextControl.ResetContents method to create a new document.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub New_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Check whether the document is dirty. In this case, the user is suggested to save that document. 
            If SaveDirtyDocumentOnNew() Then
                ' Create a new document.
                Me.m_txTextControl.ResetContents()

                ' A new document is created. Now:
                UpdateCurrentDocumentInfo() ' Reset the current document information.
                UpdateMainWindowCaption() ' Update the caption of the application's main window.
                UpdateSaveEnabledState() ' Update the enabled state of the Save button.
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SaveDirtyDocumentOnOpen Method
        ' If the document is dirty (unsaved changes were made), a MessageBox is shown where the user can
        ' decide whether loading the new file should be canceled or the changed document should (or should not) 
        ' be saved before creating a new document.
        '
        ' Return value:    If creating the new document should be canceled, the method returns false. 
        '                  Otherwise true.
        '-----------------------------------------------------------------------------------------------------------
        Private Function SaveDirtyDocumentOnNew() As Boolean
            Dim bKeepGoing = True
            If m_bIsDirtyDocument Then
                ' If the document is dirty, show a message box where the user can decide how to handle it.

                ' The message box' text depends on whether the dirty document is an unsaved file or not.
                Dim strMessageBoxText = If(m_bIsUnknownDocument, My.Resources.MessageBox_SaveDirtyDocumentOnNew_Untitled, String.Format(My.Resources.MessageBox_SaveDirtyDocumentOnNew_ToFile, m_strActiveDocumentPath))

                ' Show message box.
                Dim mbrSaveFile = MessageBox.Show(Me, strMessageBoxText, My.Resources.MessageBox_SaveDirtyDocumentOnNew_Caption, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)
                Select Case mbrSaveFile
                    Case MessageBoxResult.Yes
                        ' The dirty document should be saved.
                        bKeepGoing = Save(m_strActiveDocumentPath)
                    Case MessageBoxResult.Cancel
                        ' Opening a new document is canceled.
                        bKeepGoing = False
                End Select
            End If
            Return bKeepGoing
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateCurrentDocumentInfo Method
        ' Updates some variables that provide information about the current active document. In this case 
        ' these information are reset to the values of a newly created document.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateCurrentDocumentInfo()
            m_strActiveDocumentPath = Nothing
            m_strUserPasswordPDF = String.Empty
            m_strCssFileName = Nothing
            m_svCssSaveMode = CssSaveMode.None
            m_bIsDirtyDocument = False
            m_bIsUnknownDocument = True
            m_strActiveDocumentName = My.Resources.MainWindow_Caption_Untitled
        End Sub
    End Class
End Namespace
