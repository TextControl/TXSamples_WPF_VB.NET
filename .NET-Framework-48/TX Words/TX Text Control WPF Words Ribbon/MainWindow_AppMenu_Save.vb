'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_Save.vb File
'
' Description:
'      Handles saving the current document.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.IO

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private m_stLastSavedType As StreamType = StreamType.RichTextFormat ' The StreamType that was last used to save a document. If no document has been saved so far, RichtTextFormat is used. 


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Save_Click Handler
        ' Invokes the Save method to save a document by saving it at the same location where it was loaded 
        ' before.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Save_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Save(m_strActiveDocumentPath)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SaveAs_Click Handler
        ' Invokes the Save method to save a document by using the internal TextControl 'Save' dialog.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SaveAs_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Save(Nothing)
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Save Method
        ' Saves the current document as a file by using the TextControl Save dialog or using the path where the
        ' the active document was loaded from.
        '
        ' Parameters:
        '              savePath: The path where to save the active document. If that parameter is null  
        '                        or an empty string, the TextControl Save dialog is opened to save the 
        '                        document.
        '
        ' Return value:    False if the document was not saved. Otherwise true.
        '-----------------------------------------------------------------------------------------------------------
        Private Function Save(ByVal savePath As String) As Boolean
            ' Create settings to determine some save parameters and get information about the document
            ' when it is saved.
            Dim svsSaveSettings As SaveSettings = CreateSaveSettings()

            ' Check whether a file path is specified where the document should be loaded.
            If String.IsNullOrEmpty(savePath) Then
                ' If no such path Is determined, the TextControl Save dialog Is opened. In that dialog 
                ' all file formats can be chosen that are provided by the TXTextControl.StreamType enumeration.
                ' Furthermore the DialogSettings EnterPassword, StylesheetOptions And SaveSelection are set.
                If m_txTextControl.Save(StreamType.All, svsSaveSettings,
                                        SaveSettings.DialogSettings.EnterPassword Or
                                        SaveSettings.DialogSettings.StylesheetOptions Or
                                        SaveSettings.DialogSettings.SaveSelection) = WPF.DialogResult.Cancel Then
                    Return False
                End If
            Else
                ' Save the document at the same location (and with the same format) where it was loaded
                ' before.
                svsSaveSettings.CssSaveMode = m_svCssSaveMode ' Set the stored css save mode.
                svsSaveSettings.CssFileName = m_strCssFileName ' Set the stored css file name.
                svsSaveSettings.UserPassword = m_strUserPasswordPDF ' Set the stored user password.
                Me.m_txTextControl.Save(m_strActiveDocumentPath, m_stActiveDocumentType, svsSaveSettings)
            End If

            ' The document is saved. Now:
            UpdateCurrentDocumentInfo(svsSaveSettings)  ' Set information about the saved document.       
            AddRecentFile(m_strActiveDocumentPath) ' Add the document to the recent files list.
            UpdateMainWindowCaption()  ' Update the caption of the application's main window.
            UpdateSaveEnabledState() ' Update the enabled state of the Save button.
            Return True
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' CreateSaveSettings Method
        ' Creates and returns an object of type TXTextControl.SaveSettings that is used to save a document and
        ' provide information about the document after it was saved.
        '
        ' Return value:    The created SaveSettings object.
        '-----------------------------------------------------------------------------------------------------------
        Private Function CreateSaveSettings() As SaveSettings
            Dim svsSaveSettings As SaveSettings = New SaveSettings With {
                .LastModificationDate = Date.Now,
                .ReportingMergeBlockFormat = ReportingMergeBlockFormat.SubTextParts,
                .DefaultStreamType = m_stLastSavedType
            }
            Return svsSaveSettings
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateCurrentDocumentInfo Method
        ' Updates some variables that provide information about the current active document. In this case 
        ' these information are updated by the save settings of the saved document.
        '
        ' Parameters:
        '              saveSettings:   The save settings that provide the information about the saved 
        '                              document.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateCurrentDocumentInfo(ByVal saveSettings As SaveSettings)
            m_strActiveDocumentPath = saveSettings.SavedFile
            m_stLastSavedType = saveSettings.SavedStreamType
            m_stActiveDocumentType = m_stLastSavedType
            m_strUserPasswordPDF = saveSettings.UserPassword
            m_strCssFileName = saveSettings.CssFileName
            m_svCssSaveMode = saveSettings.CssSaveMode
            m_bIsDirtyDocument = False
            m_bIsUnknownDocument = False
            m_strActiveDocumentName = Path.GetFileName(m_strActiveDocumentPath)
        End Sub
    End Class
End Namespace
