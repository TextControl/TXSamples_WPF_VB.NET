'-----------------------------------------------------------------------------------------------------------
' MainWindow_Reporting_Behavior.vb File
'
' Description:
'		Handles the behavior of the added application menu's sample template button, the RibbonReportingTab's 
'		'Database Sample' button, 'Merge' group and the reporting "Result" tab.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Windows.Controls.Ribbon
Imports TXTextControl.DocumentServer.DataSources

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private m_bIsResultTabVisible As Boolean = False
        Private m_bShowRightSidebar As Boolean = False
        Private m_bShowLeftSidebar As Boolean = False
        Private m_bShowBottomSidebar As Boolean = False
        Private m_rtLastSelectedTab As RibbonTab = Nothing

        Private m_iCurrentMergedFile As Integer = 0
        Private m_lstMergedFiles As IList(Of Byte()) = Nothing

        Private m_rbtTemplate As Byte() = Nothing
        Private m_emEditMode_Reporting As EditMode = EditMode.Edit
        Private m_bIsReadOnly As Boolean = False
        Private m_bAllowAllowEditingFormFields As Boolean = True
        Private m_bAllowCopy As Boolean = True
        Private m_bAllowPrinting As Boolean = True
        Private m_dlgMergeSettingsDialog As MergeSettingsDialog
        Private m_dlgExportMergeResult As ExportMergeResultDialog


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------


        '-----------------------------------------------------------------------------------------------------------
        ' Application Menu
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SampleInvoiceTemplate_Click Handler
        ' Loads the sample 'Invoice' template.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SampleInvoiceTemplate_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Open(m_strFilesDirectory & "invoice.docx", StreamType.WordprocessingML)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SamplePackingListTemplate_Click Handler
        ' Loads the sample 'Packing List' template.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SamplePackingListTemplate_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Open(m_strFilesDirectory & "packinglist.docx", StreamType.WordprocessingML)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SampleShippingLabelTemplate_Click Handler
        ' Loads the sample 'Shipping Label' template.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SampleShippingLabelTemplate_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Open(m_strFilesDirectory & "shippinglabel.docx", StreamType.WordprocessingML)
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' RibbonReportingTab
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SampleDatabaseButton_Click Handler
        ' Loads the sample database.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SampleDatabaseButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_rtRibbonReportingTab.DataSourceManager.LoadXmlFile(m_strFilesDirectory & "sample_db.xml")
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' DataSourceManager_IsMergingPossibleChanged Handler
        ' Enable/Disable the 'Merge' group when the IsMergingPossible property value of the RibbonReportingTab's  
        ' DataSourceManager changed.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub DataSourceManager_IsMergingPossibleChanged(ByVal sender As Object, ByVal e As EventArgs)
            m_rgMerge.IsEnabled = Me.m_rtRibbonReportingTab.DataSourceManager.IsMergingPossible
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' MergeAndExport_Click Handler
        ' Opens an instance of the merge settings dialog to determine the settings for the following merge process. 
        ' When clicking the OK button, an instance of the MergeAndWaitDialog dialog is being opened to create the  
        ' corresponding files. When the merge process is finished, the dialog will be closed, the first created   
        ' record is loaded inside the TextControl and the "Results" tab is shown.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub MergeAndExport_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Open an instance of the merge settings dialog. 
            m_dlgMergeSettingsDialog = New MergeSettingsDialog()
            If m_dlgMergeSettingsDialog.ShowDialog().Value Then
                ' Save the current content of the TextControl
                Me.m_txTextControl.Save(m_rbtTemplate, BinaryStreamType.InternalUnicodeFormat)
                ' Create merge settings and open an instance of the MergeAndWaitDialog dialog to create the 
                ' corresponding files.
                Dim mergeSettings As MergeSettings = New MergeSettings() With {
                    .AsSingleDocument = m_dlgMergeSettingsDialog.MergeIntoSingleFile,
                    .RemoveEmptyBlocks = m_dlgMergeSettingsDialog.RemoveEmptyBlocks,
                    .RemoveEmptyLines = m_dlgMergeSettingsDialog.RemoveEmptyLines,
                    .RemoveEmptyFields = m_dlgMergeSettingsDialog.RemoveEmptyFields,
                    .RemoveEmptyImages = m_dlgMergeSettingsDialog.RemoveEmptyImages,
                    .RemoveTrailingWhitespace = m_dlgMergeSettingsDialog.RemoveTrailingWhitespace
                }

                Dim dlgMergeAndWait As MergeAndWaitDialog = New MergeAndWaitDialog(m_rbtTemplate, m_dlgMergeSettingsDialog.MaxRecords, Me.m_txTextControl, mergeSettings, Me.m_rtRibbonReportingTab)
                dlgMergeAndWait.ShowDialog()

                ' When the merge process is finished and merge files are created, the first 
                ' created record is loaded inside the TextControl and the 'Result' tab is shown.
                If dlgMergeAndWait.MergedFiles IsNot Nothing Then
                    ShowResultTab(dlgMergeAndWait.MergedFiles)
                End If

                ' Throw the exception when triggered by merge process.
                If dlgMergeAndWait.Exception IsNot Nothing Then
                    Throw dlgMergeAndWait.Exception
                End If
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' 'Reporting' ContextualTabGroup and 'Result' RibbonTab
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' ExitResultTab_Click Handler
        ' Closes/hides the 'Result' tab.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub ExitResultTab_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_rbnRibbon.SelectedItem = m_rtLastSelectedTab
            CloseResultTab()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' FirstRecord_Click Handler
        ' Loads the first merged file.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub FirstRecord_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            LoadMergeFile(0)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' PreviousRecord_Click Handler
        ' Loads the previous merged file.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub PreviousRecord_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            LoadMergeFile(m_iCurrentMergedFile - 1)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' NextRecord_Click Handler
        ' Loads the next merged file.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub NextRecord_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            LoadMergeFile(m_iCurrentMergedFile + 1)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' LastRecord_Click Handler
        ' Loads the last merged file.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub LastRecord_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            LoadMergeFile(m_lstMergedFiles.Count - 1)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Reporting_RibbonTabIndexChanged Handler
        ' Closes/hides the 'Result' tab when another tab is selected.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Reporting_RibbonTabIndexChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
            If Me.m_ctgReportingResult.Visibility = Visibility.Visible AndAlso Me.m_rbnRibbon.SelectedItem IsNot Me.m_rtMergeResultsTab Then
                CloseResultTab()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Reporting_TextControl_ContentChanged Handler
        ' Closes/hides the 'Result' tab when a new document is loaded or the TextControl content was reset.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Reporting_TextControl_ContentChanged(ByVal sender As Object, ByVal e As EventArgs)
            If Me.m_ctgReportingResult.Visibility = Visibility.Visible AndAlso Me.m_rbnRibbon.SelectedItem Is Me.m_rtMergeResultsTab Then
                m_rbtTemplate = Nothing
                Me.m_rbnRibbon.SelectedItem = m_rtLastSelectedTab
                CloseResultTab()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' ExportMergeResult_Click Handler
        ' Opens a dialog to export the merge result.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub ExportMergeResult_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)

            ' Determine the directory path that is first shown when the 'Export Merge Result' dialog is opened.
            Dim strExportDirectory = My.Settings.Default.ExportDirectory
            If String.IsNullOrEmpty(strExportDirectory) Then
                strExportDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            End If
            m_dlgExportMergeResult = New ExportMergeResultDialog()
            m_dlgExportMergeResult.Directory = strExportDirectory

            ' Set the current active document name as file prefix.
            m_dlgExportMergeResult.FilePrefix = m_strActiveDocumentName

            ' Open the 'Export Merge Result' dialog
            If m_dlgExportMergeResult.ShowDialog().Value Then
                ' Get the directory where to export the merged files.
                strExportDirectory = IO.Path.GetFullPath(m_dlgExportMergeResult.Directory)

                ' Use a ServerTextControl to load and export each merged file.
                Using serverTextControl As ServerTextControl = New ServerTextControl()
                    serverTextControl.Create()
                    For i = 0 To m_lstMergedFiles.Count - 1
                        ' Load the merged file into the ServerTextControl.
                        serverTextControl.Load(m_lstMergedFiles(i), BinaryStreamType.InternalUnicodeFormat)

                        ' Export/Save the merged file from the ServerTextControl.
                        Dim strPath = strExportDirectory & "\" & m_dlgExportMergeResult.FilePrefix & i + 1 & m_dlgExportMergeResult.Format.Extension
                        serverTextControl.Save(strPath, m_dlgExportMergeResult.Format.StreamType)
                    Next
                End Using

                ' Open the directory where the merged files are exported.
                If m_dlgExportMergeResult.openDirectory Then
                    Process.Start(strExportDirectory)
                End If

                ' Store the current used export directory inside the application settings.
                My.Settings.Default.ExportDirectory = strExportDirectory
                Call My.Settings.Default.Save()
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' 'Reporting' ContextualTabGroup and 'Result' RibbonTab
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetResultTabBehavior Method
        ' Manages the 'Result' tab behavior that should be automatically closed when another tab is selected, a
        ' document is loaded or the TextControl content is reset.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetResultTabBehavior()
            AddHandler Me.m_rbnRibbon.SelectionChanged, AddressOf Me.Reporting_RibbonTabIndexChanged
            AddHandler Me.m_txTextControl.ContentsReset, AddressOf Me.Reporting_TextControl_ContentChanged
            AddHandler Me.m_txTextControl.DocumentLoaded, AddressOf Me.Reporting_TextControl_ContentChanged
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' ShowResultTab Method
        ' Shows the 'Result' tab with the first record of the specified merged files.
        '
        ' Parameters:
        '      mergedFiles:    The merged files that should be provided to show inside the 'Result' tab.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub ShowResultTab(ByVal mergedFiles As IList(Of Byte()))
            ' Hide all sidebars
            HideSidebars()

            ' Store the current TextControl document permissions and edit mode
            m_bIsReadOnly = Me.m_txTextControl.DocumentPermissions.ReadOnly
            m_bAllowAllowEditingFormFields = Me.m_txTextControl.DocumentPermissions.AllowEditingFormFields
            m_bAllowCopy = Me.m_txTextControl.DocumentPermissions.AllowCopy
            m_bAllowPrinting = Me.m_txTextControl.DocumentPermissions.AllowPrinting
            m_emEditMode_Reporting = Me.m_txTextControl.EditMode

            ' Remember the current selected tab.
            m_rtLastSelectedTab = TryCast(Me.m_rbnRibbon.SelectedItem, RibbonTab)

            ' Load the first record of the specified merged file.
            m_iCurrentMergedFile = 0
            m_lstMergedFiles = mergedFiles
            LoadMergeFile(m_iCurrentMergedFile)

            ' Show the 'Result' contextual tab.
            Me.m_ctgReportingResult.Visibility = Visibility.Visible
            ' Ensure that the "Table" contextual tab is hidden.
            Me.m_ctgTableTools.Visibility = Visibility.Collapsed
            ' Show the 'Result' Tab.
            Me.m_rbnRibbon.SelectedItem = Me.m_rtMergeResultsTab
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' CloseResultTab Method
        ' Closes/Hides the 'Result' tab.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub CloseResultTab()
            ' Reset some variables.
            m_iCurrentMergedFile = 0
            m_lstMergedFiles = Nothing
            ' Hide the 'Reporting' contextual tab.
            Me.m_ctgReportingResult.Visibility = Visibility.Collapsed

            ' Reset the content and state to what was shown before the 'Result' tab was shown.
            If m_rbtTemplate IsNot Nothing Then
                Me.m_txTextControl.Load(m_rbtTemplate, BinaryStreamType.InternalUnicodeFormat)
            End If
            ResetSaveButtonBehavior()

            ' Set the edit mode and document permission to what were set before the 'Result' tab was shown.
            Me.m_txTextControl.DocumentPermissions.ReadOnly = m_bIsReadOnly
            Me.m_txTextControl.DocumentPermissions.AllowEditingFormFields = m_bAllowAllowEditingFormFields
            Me.m_txTextControl.DocumentPermissions.AllowCopy = m_bAllowCopy
            Me.m_txTextControl.DocumentPermissions.AllowPrinting = m_bAllowPrinting
            Me.m_txTextControl.EditMode = m_emEditMode_Reporting

            ' Show the sidebars that where shown before the 'Result' tab was shown.
            ShowSidebars()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' LoadMergeFile Method
        ' Load the merged file that is stored at the specified index inside merged files list.
        '
        ' Parameters:
        '      index:		The index of the merged file to load.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub LoadMergeFile(ByVal index As Integer)
            '  Load the merged file that is stored at the specified index inside merged files list.
            m_iCurrentMergedFile = index
            RemoveHandler Me.m_txTextControl.DocumentLoaded, AddressOf Me.Reporting_TextControl_ContentChanged ' Disable DocumentLoaded event handling
            Me.m_txTextControl.Load(m_lstMergedFiles(m_iCurrentMergedFile), BinaryStreamType.InternalUnicodeFormat) ' Load the merged file.
            AddHandler Me.m_txTextControl.DocumentLoaded, AddressOf Me.Reporting_TextControl_ContentChanged ' Enable DocumentLoaded event handling
            ' The loaded merged file should not be editable.
            Me.m_txTextControl.DocumentPermissions.ReadOnly = True
            Me.m_txTextControl.DocumentPermissions.AllowEditingFormFields = True
            Me.m_txTextControl.DocumentPermissions.AllowCopy = True
            Me.m_txTextControl.DocumentPermissions.AllowPrinting = True
            Me.m_txTextControl.EditMode = EditMode.ReadAndSelect

            ' Update enabled states
            Me.m_rbtnFirstRecord.IsEnabled = CSharpImpl.Assign(Me.m_rbtnPreviousRecord.IsEnabled, m_iCurrentMergedFile > 0)
            Me.m_rbtnNextRecord.IsEnabled = CSharpImpl.Assign(Me.m_rbtnLastRecord.IsEnabled, m_iCurrentMergedFile < m_lstMergedFiles.Count - 1)

            DisableSaveButtonBehavior()

            ' Update tool tip despriptions
            Dim strToolTipDescription = My.Resources.MergeResultsTab_Record_ToolTip_Description
            Me.m_rbtnFirstRecord.ToolTipDescription = String.Format(strToolTipDescription, 1)
            Me.m_rbtnPreviousRecord.ToolTipDescription = String.Format(strToolTipDescription, m_iCurrentMergedFile)
            Me.m_rbtnNextRecord.ToolTipDescription = String.Format(strToolTipDescription, m_iCurrentMergedFile + 2)
            Me.m_rbtnLastRecord.ToolTipDescription = String.Format(strToolTipDescription, m_lstMergedFiles.Count)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' HideSidebars Method
        ' Stores the Sidebar's IsShown property value of each sidebar before hiding them.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub HideSidebars()
            If m_bIsResultTabVisible <> CSharpImpl.Assign(m_bIsResultTabVisible, True) Then
                ' Store the Sidebar's IsShown property value of each sidebar.
                m_bShowRightSidebar = Me.m_sbSidebarRight.IsShown
                m_bShowLeftSidebar = Me.m_sbSidebarLeft.IsShown
                m_bShowBottomSidebar = Me.m_sbSidebarBottom.IsShown
                ' Hide sidebars
                Me.m_sbSidebarRight.IsShown = CSharpImpl.Assign(Me.m_sbSidebarLeft.IsShown, CSharpImpl.Assign(Me.m_sbSidebarBottom.IsShown, False))
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' ShowSidebars Method
        ' Shows the sidebars that where shown before displaying the 'Result' tab.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub ShowSidebars()
            If m_bIsResultTabVisible <> CSharpImpl.Assign(m_bIsResultTabVisible, False) Then
                Me.m_sbSidebarRight.IsShown = m_bShowRightSidebar
                Me.m_sbSidebarLeft.IsShown = m_bShowLeftSidebar
                Me.m_sbSidebarBottom.IsShown = m_bShowBottomSidebar
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' DisableSaveButtonBehavior Method
        ' Ensures that the template cannot be overridden by using the 'Save' button as long as the 'Result' tab is 
        ' shown.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub DisableSaveButtonBehavior()
            RemoveHandler Me.m_txTextControl.Changed, AddressOf Me.TextControl_Changed
            m_bIsDirtyDocument = False
            UpdateMainWindowCaption()
            UpdateSaveEnabledState()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' ResetSaveButtonBehavior Method
        ' Resets the 'Save' button behavior.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub ResetSaveButtonBehavior()
            AddHandler Me.m_txTextControl.Changed, AddressOf Me.TextControl_Changed
        End Sub
    End Class
End Namespace
