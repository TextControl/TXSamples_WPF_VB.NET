'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_RecentFiles.vb File
'
' Description:
'     Manages the recent files.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Collections.Specialized
Imports System.IO
Imports System.Windows.Controls.Ribbon

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private Const m_iMaxRecentFiles As Integer = 10
        Private m_colRecentFiles As StringCollection


        '-----------------------------------------------------------------------------------------------------------
        ' P R O P E R T I E S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' RecentFiles Property
        ' Gets or sets the list of recent files.
        '-----------------------------------------------------------------------------------------------------------
        Friend Property RecentFiles As StringCollection
            Get
                Return m_colRecentFiles
            End Get
            Set(ByVal value As StringCollection)
                m_colRecentFiles = If(value, New StringCollection())
                ' Remove empty entries.
                For i = m_colRecentFiles.Count - 1 To 0 Step -1
                    If String.IsNullOrEmpty(m_colRecentFiles(i)) Then
                        m_colRecentFiles.RemoveAt(i)
                    End If
                Next
                UpdateRecentFileList()
            End Set
        End Property


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' RecentFiles_SelectionChanged Handler
        ' Opens the file that is represented by the selected RibbonGalleryItem.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub RecentFiles_SelectionChanged(ByVal sender As Object, ByVal e As RoutedPropertyChangedEventArgs(Of Object))
            ' Get the file that is represented by the selected RibbonGalleryItem.
            Dim strFile As String = TryCast(TryCast(e.NewValue, RibbonGalleryItem).Tag, String)

            ' Check whether the file still exists.
            If Not File.Exists(strFile) Then
                ' If not, open a Message box that asks the user whether the not
                ' existing file should be removed from the recent file list.
                If MessageBox.Show(Me, My.Resources.MessageBox_OpenRecentFile_FileDoesNotExist_Text, My.Resources.MessageBox_OpenRecentFile_FileDoesNotExist_Caption, MessageBoxButton.OKCancel, MessageBoxImage.Question) = MessageBoxResult.OK Then
                    ' Remove the file from the list and update the items collection.
                    m_colRecentFiles.Remove(strFile)
                    UpdateRecentFileList()
                End If
            Else
                ' Open the file.
                Open(strFile)
            End If

            Me.m_rgalRecentFiles.SelectedItem = Nothing ' Deselect recent file.
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' LoadRecentFiles Method
        ' Gets the recent files from the application settings.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub LoadRecentFiles()
            RecentFiles = My.Settings.Default.RecentFiles
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SaveRecentFiles Method
        ' Saves the recent files list to the My.Settings.Default.RecentFiles property when the 
        ' application is closing (see MainWindow_FormClosing Handler).
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SaveRecentFiles()
            My.Settings.Default.RecentFiles = RecentFiles
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

                ' Update the recent file items inside the ribbon's ApplicationMenuHelpPaneItems collection.
                UpdateRecentFileList()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateRecentFileList Method
        ' Updates the recent file items inside the ribbon's ApplicationMenuHelpPaneItems collection.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateRecentFileList()

            ' Remove all recent file items from the collection.
            Me.m_rgcRecentFiles.Items.Clear()

            ' Create and insert for each recent file path entry a RibbonButton that represents a recent file.
            Dim i = 0

            While i < m_colRecentFiles.Count AndAlso i < m_iMaxRecentFiles
                Me.m_rgcRecentFiles.Items.Add(CreateRecentFileButton(i))
                i += 1
            End While

        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' CreateRecentFileButton Method
        ' Creates and returns a RibbonButton that represents a recent file.
        '
        ' Parameters:
        '      index:   The index of the recent file inside the recent files collection.
        '
        ' Return value:    A RibbonButton that represents a recent file.
        '-----------------------------------------------------------------------------------------------------------
        Private Function CreateRecentFileButton(ByVal index As Integer) As RibbonGalleryItem
            ' Create a labeled RibbonButton with no icon
            Dim rbtnRecentFileButton As RibbonGalleryItem = New RibbonGalleryItem()

            ' Get the path and name of the file.
            Dim strFilePath = m_colRecentFiles(index)
            Dim strFileName = Path.GetFileName(strFilePath)

            ' Determine the displayed text of the button (index plus file name) 
            ' and store the file path as Tag value.
            rbtnRecentFileButton.KeyTip = (index + 1).ToString()
            rbtnRecentFileButton.Content = rbtnRecentFileButton.KeyTip & ". " & strFileName
            rbtnRecentFileButton.Tag = strFilePath

            ' Provide file name and path by setting the tool tip.
            rbtnRecentFileButton.ToolTipTitle = strFileName
            rbtnRecentFileButton.ToolTipDescription = strFilePath

            Return rbtnRecentFileButton
        End Function
    End Class
End Namespace
