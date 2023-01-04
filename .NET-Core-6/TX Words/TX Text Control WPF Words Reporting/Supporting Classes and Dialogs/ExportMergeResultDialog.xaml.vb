Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for ExportMergeResultDialog.xaml
    ''' </summary>
    Partial Public Class ExportMergeResultDialog
        Inherits Window
        '----------------------------------------------------------------------------------------------
        ' Class FormatItem
        ' Represents the 'Format' combo box item. It provides the displayed text, the format extension
        ' and the TXTextControl.StreamType to use.
        '----------------------------------------------------------------------------------------------
        Friend Class FormatItem
            ' Member Variables
            Private m_strFormat As String
            Private m_strExtension As String
            Private m_stStreamType As StreamType

            ' Constructor
            Friend Sub New(ByVal displayedText As String, ByVal extension As String, ByVal streamType As StreamType)
                m_strFormat = displayedText
                m_strExtension = extension
                m_stStreamType = streamType
            End Sub

            ' Properties
            Friend ReadOnly Property Extension As String
                Get
                    Return m_strExtension
                End Get
            End Property

            Friend ReadOnly Property StreamType As StreamType
                Get
                    Return m_stStreamType
                End Get
            End Property

            ' Overridden Methods.
            Public Overrides Function ToString() As String
                Return String.Format(m_strFormat, m_strExtension)
            End Function
        End Class


        '----------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R
        '----------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' ExportMergeResultDialog Constructor
        ' Creates a dialog to export the results of the merge process.
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New()
            Me.InitializeComponent()

            ' Set some texts
            Me.Title = My.Resources.ExportMergeResultDialog_Caption
            Me.m_lblFilePrefix.Content = My.Resources.ExportMergeResultDialog_FilePrefix
            Me.m_lblDirectory.Content = My.Resources.ExportMergeResultDialog_Directory
            Me.m_lblFormat.Content = My.Resources.ExportMergeResultDialog_Format
            Me.m_chbxopenDirectory.Content = My.Resources.ExportMergeResultDialog_openDirectory
            Me.m_btnOK.Content = My.Resources.ExportMergeResultDialog_OK
            Me.m_btnCancel.Content = My.Resources.ExportMergeResultDialog_Cancel

            ' Add format items.
            Me.m_cmbxFormat.Items.Add(New FormatItem(My.Resources.ExportMergeResultDialog_Format_RTF, ".rtf", StreamType.RichTextFormat))
            Me.m_cmbxFormat.Items.Add(New FormatItem(My.Resources.ExportMergeResultDialog_Format_HTML, ".html", StreamType.HTMLFormat))
            Me.m_cmbxFormat.Items.Add(New FormatItem(My.Resources.ExportMergeResultDialog_Format_DOCX, ".docx", StreamType.SpreadsheetML))
            Me.m_cmbxFormat.Items.Add(New FormatItem(My.Resources.ExportMergeResultDialog_Format_DOC, ".doc", StreamType.MSWord))
            Me.m_cmbxFormat.Items.Add(New FormatItem(My.Resources.ExportMergeResultDialog_Format_PDF, ".pdf", StreamType.AdobePDFA))
            Me.m_cmbxFormat.Items.Add(New FormatItem(My.Resources.ExportMergeResultDialog_Format_TXT, ".txt", StreamType.PlainText))
            Me.m_cmbxFormat.Items.Add(New FormatItem(My.Resources.ExportMergeResultDialog_Format_TX, ".tx", StreamType.InternalFormat))

            ' Select the PDF item.
            Me.m_cmbxFormat.SelectedIndex = 4
        End Sub


        '----------------------------------------------------------------------------------------------
        ' P R O P E R T I E S
        '----------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Directory Property
        ' Gets or sets the directoy where to export the created files of the merge process. 
        '-----------------------------------------------------------------------------------------------------------
        Friend Property Directory As String
            Get
                Return Me.m_tbxDirectory.Text.Trim()
            End Get
            Set(ByVal value As String)
                Me.m_tbxDirectory.Text = value
            End Set
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' FilePrefix Property
        ' Gets or sets the prefix string that is used when exporting the created files of the merge process.
        '-----------------------------------------------------------------------------------------------------------
        Friend Property FilePrefix As String
            Get
                Return Me.m_tbxFilePrefix.Text.Trim()
            End Get
            Set(ByVal value As String)
                Me.m_tbxFilePrefix.Text = IO.Path.GetFileNameWithoutExtension(value)
            End Set
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' Format Property
        ' Gets the document format that is used when exporting the created files of the merge process.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property Format As FormatItem
            Get
                Return TryCast(Me.m_cmbxFormat.SelectedItem, FormatItem)
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' openDirectory Property
        ' Gets a value indicating whether the directory where the merged files are exported should be openeded.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property openDirectory As Boolean
            Get
                Return Me.m_chbxopenDirectory.IsChecked.Value
            End Get
        End Property


        '----------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '----------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Directory_TextChanged Handler
        ' Handles the enabled state of the OK button.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Directory_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
            Me.m_btnOK.IsEnabled = Directory.Length > 0
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' OK_Click Handler
        ' If the specified directory path exists, close the dialog.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If Not IO.Directory.Exists(Directory) Then
                MessageBox.Show(My.Resources.MessageBox_ExportMergeResultDialog_DirectoryDoesNotExist_Text, My.Resources.MessageBox_ExportMergeResultDialog_DirectoryDoesNotExist_Caption, MessageBoxButton.OK, MessageBoxImage.Error)
            Else
                DialogResult = True
                Close()
            End If
        End Sub
    End Class
End Namespace
