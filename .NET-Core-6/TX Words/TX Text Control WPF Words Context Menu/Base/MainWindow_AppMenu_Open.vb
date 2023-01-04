'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_Open.vb File
'
' Description:
'     Manages opening a file.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.IO

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private m_stLastLoadedType As StreamType = StreamType.RichTextFormat
        Private ReadOnly m_isPDFImportSettings As PDFImportSettings = PDFImportSettings.GenerateTextFrames Or PDFImportSettings.LoadEmbeddedFiles


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Open_Click Handler
        ' Invokes the Open method to load a document by using the internal TextControl 'Open' dialog.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Open_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Open()
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Open Method
        ' Opens a file.
        '
        ' Parameters:
        '      fileName:   If not specified, the TextControl Load-Dialog is opened to select the file to open. 
        '                  Otherwise the specified file is loaded.
        '      streamType: If set, the TXTextControl.StreamType value is used to load the file that is specified
        '                  by the fileName parameter. Otherwise, the corresponding StreamType is determined by
        '                  the file's format.
        '-----------------------------------------------------------------------------------------------------------
        Friend Sub Open(ByVal Optional fileName As String = Nothing, ByVal Optional streamType As StreamType = -1)
            ' Check whether the document is dirty. In this case, the user is suggested to save that document. 
            If SaveDirtyDocumentOnOpen() Then
                ' Create settings to determine some load parameters and get information about the document
                ' when it is opened.
                Dim lsLoadSettings As LoadSettings = CreateLoadSettings()
                Try
                    ' Check whether a file to load is specified.
                    If String.IsNullOrEmpty(fileName) Then
                        ' If Not, the TextControl Load dialog Is opened. In that dialog all loadable file 
                        ' formats can be chosen that are provided by the TXTextControl.StreamType enumeration.
                        If Me.m_txTextControl.Load(StreamType.All, lsLoadSettings) = WPF.DialogResult.Cancel Then
                            Return
                        End If
                    Else
                        ' Determine the stream type if necessary
                        If streamType = CType(-1, StreamType) Then
                            streamType = GetStreamType(fileName)
                        End If
                        ' Open the file directly by using its path.
                        Me.m_txTextControl.Load(fileName, streamType, lsLoadSettings)
                    End If
                Catch ex As Exception
                    ' Set the password if the document is password protected.
                    If Not HandlePasswordProtectedDocument(ex, lsLoadSettings) Then
                        Return
                    End If
                End Try
                ' The document is loaded. Now:
                UpdateCurrentDocumentInfo(lsLoadSettings)  ' Set information about the loaded document.              
                AddRecentFile(m_strActiveDocumentPath)  ' Add the document to the recent files list.
                UpdateMainWindowCaption() ' Update the caption of the application's main window.
                UpdateSaveEnabledState() ' Update the enabled state of the Save button.
            End If
        End Sub

        ' -------------------------------------------------------------------------------------------------------------
        ' HandlePasswordProtectedDocument Method
        ' Handles password protect documents by opening a password dialog.
        '
        ' Parameters:
        ' 		exception:		The exception that is thrown when opening the document.
        ' 		loadSettings:	The load settings that are used when opening the doucment.
        ' 
        ' 		Return value:		True if the password protected document could be loaded.
        '					        Otherwise false.
        '-----------------------------------------------------------------------------------------------------------
        Private Function HandlePasswordProtectedDocument(ByVal exception As Exception, ByVal loadSettings As LoadSettings) As Boolean
            ' Check whether the thrown exception is an exception of type FilterException.
            If TypeOf exception Is FilterException Then
                Select Case TryCast(exception, FilterException).Reason
                    Case FilterException.FilterError.InvalidPassword
                        ' Open the password dialog if the document is write protected.
                        Dim dlgPassword As PasswordDialog = New PasswordDialog(Me.m_txTextControl, loadSettings)
                        dlgPassword.Owner = Me
                        Dim bResult As Boolean? = dlgPassword.ShowDialog()
                        Return bResult.HasValue AndAlso bResult.Value
                End Select
            End If
            Throw exception
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' SaveDirtyDocumentOnOpen Method
        ' If the document is dirty (unsaved changes were made), a MessageBox is shown where the user can
        ' decide whether loading the new file should be canceled or the changed document should (or should not) 
        ' be saved before opening the new document.
        '
        ' Return value:    If loading the new document should be canceled, the method returns false. 
        '                  Otherwise true.
        '-----------------------------------------------------------------------------------------------------------
        Private Function SaveDirtyDocumentOnOpen() As Boolean
            Dim bKeepGoing = True
            If m_bIsDirtyDocument Then
                ' If the document is dirty, show a message box where the user can decide how to handle it.

                ' The message box' text depends on whether the dirty document is an unsaved file or not.
                Dim strMessageBoxTExt = If(m_bIsUnknownDocument, My.Resources.MessageBox_SaveDirtyDocumentOnOpen_Untitled, String.Format(My.Resources.MessageBox_SaveDirtyDocumentOnOpen_ToFile, m_strActiveDocumentPath))

                ' Show message box.
                Dim mbrSaveFile = MessageBox.Show(Me, strMessageBoxTExt, My.Resources.MessageBox_SaveDirtyDocumentOnOpen_Caption, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)

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
        ' GetStreamType Method
        ' Checks whether the format of the specified file is supported by TX Text Control and returns the
        ' corresponding StreamType.
        '
        ' Parameters:
        '      filePath:   The file to determine whether or not its format is supported by TX Text Control.
        '
        ' Return value:    The corresponding StreamType if the file format is supported. Otherwise 
        '                  (StreamType)(-1)
        '-----------------------------------------------------------------------------------------------------------
        Private Function GetStreamType(ByVal filePath As String) As StreamType
            Dim strFileExtension = Path.GetExtension(filePath)
            Select Case strFileExtension
                Case ".rtf"
                    Return StreamType.RichTextFormat

                Case ".htm", ".html"
                    Return StreamType.HTMLFormat

                Case ".tx"
                    Return StreamType.InternalUnicodeFormat

                Case ".doc"
                    Return StreamType.MSWord

                Case ".docx"
                    Return StreamType.WordprocessingML

                Case ".pdf"
                    Return StreamType.AdobePDF

                Case ".txt"
                    Return StreamType.PlainText

                Case ".xlsx"
                    Return StreamType.SpreadsheetML
            End Select
            Return -1
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' CreateLoadSettings Method
        ' Creates and returns an object of type TXTextControl.LoadSettings that is used to open a document and
        ' provide information about the document after it was loaded.
        '
        ' Return value:    The created LoadSettings object.
        '-----------------------------------------------------------------------------------------------------------
        Private Function CreateLoadSettings() As LoadSettings
            Dim lsLoadSettings As LoadSettings = New LoadSettings With {
                .ApplicationFieldFormat = ApplicationFieldFormat.MSWordTXFormFields,
                .LoadSubTextParts = True,
                .ReportingMergeBlockFormat = ReportingMergeBlockFormat.SubTextParts,
                .PDFImportSettings = m_isPDFImportSettings,
                .DocumentPartName = String.Empty,
                .DefaultStreamType = m_stLastLoadedType
            }
            Return lsLoadSettings
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateCurrentDocumentInfo Method
        ' Updates some variables that provide information about the current active document. In this case 
        ' these information are updated by the load settings of the opened document.
        '
        ' Parameters:
        '              loadSettings:   The load settings that provide the information about the opened 
        '                              document.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateCurrentDocumentInfo(ByVal loadSettings As LoadSettings)
            m_strActiveDocumentPath = loadSettings.LoadedFile
            m_stLastLoadedType = loadSettings.LoadedStreamType
            m_stActiveDocumentType = m_stLastLoadedType
            m_strUserPasswordPDF = loadSettings.UserPassword
            m_strCssFileName = loadSettings.CssFileName
            m_svCssSaveMode = CssSaveMode.None
            m_bIsDirtyDocument = False
            m_bIsUnknownDocument = False
            m_strActiveDocumentName = Path.GetFileName(m_strActiveDocumentPath)
        End Sub
    End Class
End Namespace
