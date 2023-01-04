'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_FileInfo.vb File
'
' Description:
'     Manages updating information about the current loaded/shown document.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------

        ' Values that are updated when opening, creating or saving a document
        Private m_strActiveDocumentName As String = My.Resources.MainWindow_Caption_Untitled ' The document's name is '[Untitled]' on default.
        Private m_strActiveDocumentPath As String = Nothing ' The path of the active document.
        Private m_stActiveDocumentType As StreamType = StreamType.RichTextFormat ' The StreamType that was last used To load Or save the current document.
        Private m_strUserPasswordPDF As String = String.Empty ' Tthe password for the user when the document is reopened.
        Private m_strCssFileName As String = Nothing 'The path and filename of a CSS file belonging to a HTML document.
        Private m_svCssSaveMode As CssSaveMode = CssSaveMode.None ' Specifies how to save stylesheet data with a HTML document.
        Private m_bIsUnknownDocument As Boolean = True ' A flag that indicates whether or not the active document is loaded/saved or created (unknown).

        ' A flag that indicates whether or not the document is 'dirty'
        Private m_bIsDirtyDocument As Boolean = False


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_Changed Handler
        ' Updates the 'Is Dirty Document' flag to true if the document changed.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_Changed(ByVal sender As Object, ByVal e As EventArgs)
            If m_bIsDirtyDocument <> CSharpImpl.Assign(m_bIsDirtyDocument, True) Then
                ' Update caption and save button enabled state.
                UpdateMainWindowCaption()
                UpdateSaveEnabledState()
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateMainWindowCaption Method
        ' Updates the application caption with the name of the active document and the product name.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateMainWindowCaption()
            Title = m_strActiveDocumentName & If(m_bIsDirtyDocument, "*", "") & " - " & My.Resources.MainWindow_Caption_Product
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateSaveEnabledState Method
        ' Enables the Save button in case the loaded document is dirty. Otherwise the button is disabled.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateSaveEnabledState()
            Me.m_rbtnSaveQAT.IsEnabled = CSharpImpl.Assign(Me.m_rmiSave.IsEnabled, m_bIsDirtyDocument AndAlso Not m_bIsUnknownDocument)
        End Sub
    End Class
End Namespace
