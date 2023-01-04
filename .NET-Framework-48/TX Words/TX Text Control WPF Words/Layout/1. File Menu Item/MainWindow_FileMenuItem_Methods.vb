'-----------------------------------------------------------------------------------------------------------
' MainWindow_FileMenuItem_Methods.vb File
'
' Description: Provides supporting methods to implement the desired behavior of some 'File' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Imports System.IO

Namespace TXTextControl.Words
    Partial Class MainWindow


        '-----------------------------------------------------------------------------------------------------------
        ' 'New' item
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
                Dim mrSaveFile = MessageBox.Show(Me, strMessageBoxText, My.Resources.MessageBox_SaveDirtyDocumentOnNew_Caption, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)

                Select Case mrSaveFile
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


        '-----------------------------------------------------------------------------------------------------------
        ' 'Open...' item
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
        Private Sub Open(ByVal Optional fileName As String = Nothing, ByVal Optional streamType As StreamType = -1)
            ' Check whether the document is dirty. In this case, the user is suggested to save that document. 
            If SaveDirtyDocumentOnOpen() Then
                ' Create settings to determine some load parameters and get information about the document
                ' when it is opened.
                Dim lsLoadSettings As LoadSettings = CreateLoadSettings()
                ' Check whether a file to load is specified.
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
                Dim mrSaveFile = MessageBox.Show(Me, strMessageBoxTExt, My.Resources.MessageBox_SaveDirtyDocumentOnOpen_Caption, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)

                Select Case mrSaveFile
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


        '-----------------------------------------------------------------------------------------------------------
        ' 'Recent Files' item
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' LoadRecentFiles Method
        ' Gets the recent files from the application settings.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub LoadRecentFiles()
            m_colRecentFiles = My.Settings.Default.RecentFiles

            If m_colRecentFiles Is Nothing Then
                m_colRecentFiles = New Specialized.StringCollection()
            End If
            ' Remove empty entries.
            For i = m_colRecentFiles.Count - 1 To 0 Step -1

                If String.IsNullOrEmpty(m_colRecentFiles(i)) Then
                    m_colRecentFiles.RemoveAt(i)
                End If
            Next

            UpdateRecentFileList()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SaveRecentFiles Method
        ' Saves the recent files list to the My.Settings.Default.RecentFiles property when the 
        ' application is closing (see MainWindow_FormClosing Handler).
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SaveRecentFiles()
            My.Settings.Default.RecentFiles = m_colRecentFiles
            Call My.Settings.Default.Save()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' AddRecentFile Method
        ' Inserts the specified file path as first entry inside the recent files list. 
        '
        ' Parameters:
        '      filePath:   The file path to add.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub AddRecentFile(ByVal filePath As String)
            If Not String.IsNullOrEmpty(filePath) Then
                ' Check whether the list already contains that file.
                If m_colRecentFiles.Contains(filePath) Then
                    ' In that case remove that file.
                    m_colRecentFiles.Remove(filePath)
                Else
                    ' Remove last entry if the current number of entries equals to the
                    ' maximum number of recent files.
                    If m_colRecentFiles.Count = m_iMaxRecentFiles Then
                        m_colRecentFiles.RemoveAt(m_iMaxRecentFiles - 1)
                    End If
                End If
                ' Insert the file path at the top of the list.
                m_colRecentFiles.Insert(0, filePath)

                ' Update the recent file drop down items.
                UpdateRecentFileList()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateRecentFileList Method
        ' Updates the recent file drop down items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateRecentFileList()
            Me.m_miFile_RecentFiles.Items.Clear()

            ' Create and insert for each recent file path entry an item that represents a recent file.
            Dim i = 0

            While i < m_colRecentFiles.Count AndAlso i < m_iMaxRecentFiles
                Me.m_miFile_RecentFiles.Items.Add(CreateRecentFileItem(i))
                i += 1
            End While
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' CreateRecentFileItem Method
        ' Creates and returns an item that represents a recent file.
        '
        ' Parameters:
        '      index:   The index of the recent file inside the recent files collection.
        '
        ' Return value:    A MenuItem that represents a recent file.
        '-----------------------------------------------------------------------------------------------------------
        Private Function CreateRecentFileItem(ByVal index As Integer) As MenuItem
            ' Create an item
            Dim tmiRecentFile As MenuItem = New MenuItem()

            ' Get the path and name of the file.
            Dim strFilePath = m_colRecentFiles(index)
            Dim strFileName = Path.GetFileName(strFilePath)

            ' Determine the displayed text of the item (index plus file name) 
            ' and store the file path as Tag value.
            tmiRecentFile.Header = "_" & index + 1 & ". " & strFileName
            tmiRecentFile.Tag = strFilePath

            ' Provide file path by setting the tool tip.
            tmiRecentFile.ToolTip = strFilePath

            ' Add a handler to the Click event to open the corresponding file when the item is clicked.
            AddHandler tmiRecentFile.Click, AddressOf File_RecentFiles_Item_Click
            Return tmiRecentFile
        End Function


        '-----------------------------------------------------------------------------------------------------------
        ' 'Save' and 'Save as...' items
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
                                        SaveSettings.DialogSettings.SaveSelection) = WPF.DialogResult.OK Then
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


        '-----------------------------------------------------------------------------------------------------------
        ' 'Sign In...' and '[Current user]' items
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' LoadKnownUserSettings Method
        ' Gets the known user from the application settings.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub LoadKnownUserSettings()
            m_strUserName = Environment.UserName  ' Get the user name of the person who is currently logged on the operation system
            m_uiCurrentUser = My.Settings.Default.KnownUser
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SaveKnownUserSettings Method
        ' Save the known user to the application settings.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SaveKnownUserSettings()

            ' Save the know users to the My.Settings.Default.KnownUsers property
            My.Settings.Default.KnownUser = m_uiCurrentUser
            Call My.Settings.Default.Save()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SignOut Method
        ' Signs out the current user.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SignOut()
            'Signs out the current user.
            If m_uiCurrentUser IsNot Nothing Then
                m_uiCurrentUser.IsSignedIn = False
            End If
            ' Reset the TextControl user access.
            Me.m_txTextControl.UserNames = Nothing

            ' Show the Sign In item.
            Me.m_miFile_SignIn.IsEnabled = True
            Me.m_miFile_SignIn.Visibility = Visibility.Visible

            ' Hide the [Current User] item.
            Me.m_miFile_CurrentUser.IsEnabled = False
            Me.m_miFile_CurrentUser.Visibility = Visibility.Collapsed
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' 'Exit' item
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
                Dim mrSaveFile = MessageBox.Show(Me, strMessageBoxTExt, My.Resources.MessageBox_SaveDirtyDocumentOnExit_Caption, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)

                Select Case mrSaveFile
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
