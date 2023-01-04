'-----------------------------------------------------------------------------------------------------------
' MainWindow_OpenHyperlink.vb File
'
' Description:
'     Handles opening a hyperlink
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.IO
Imports System.Reflection

Namespace TXTextControl.Words
    Partial Class MainWindow
        ' -------------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        ' -----------------------------------------------------------------------------------------------------------
        Private m_strLinkedFile As String = Nothing


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_HypertextLinkClicked Handler
        ' Open the hypertextlink.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_HypertextLinkClicked(ByVal sender As Object, ByVal e As HypertextLinkEventArgs)
            OpenHyperlink(e.HypertextLink.Target)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' MainWindow_OpenHyperlink_Load Handler
        ' If the clicked link is a document file, the link is added as Process.StartInfo argument of a new instance
        ' of this application where the file is loaded.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub MainWindow_OpenHyperlink_Load(ByVal sender As Object, ByVal e As EventArgs)
            ' Load file provided as a command line argument
            Dim rstrArgs As String() = Environment.GetCommandLineArgs()
            If rstrArgs.Length > 1 Then
                Dim strFile = rstrArgs(1)
                If File.Exists(strFile) Then
                    m_strLinkedFile = strFile
                    AddHandler Me.m_txTextControl.Loaded, AddressOf Me.TextControl_Loaded
                End If
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_DocumentLinkClicked Handler
        ' Scroll to the linked document target.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_DocumentLinkClicked(ByVal sender As Object, ByVal e As DocumentLinkEventArgs)
            Dim dtTarget As DocumentTarget = e.DocumentLink.DocumentTarget
            ' TextControl scrolls automatically to TOC targets when pressing the Ctrl key and clicking the link.
            If dtTarget IsNot Nothing AndAlso dtTarget.AutoGenerationType <> AutoGenerationType.TableOfContents Then
                dtTarget.ScrollTo()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_Loaded Handler
        ' Opens the linked file.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.Open(m_strLinkedFile)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetOpenHyperlinkBehavior Method
        ' Adds all necessary handlers to implement an open hyperlink behavior.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetOpenHyperlinkBehavior()
            ' Open Hyperlink on click:
            AddHandler Me.m_txTextControl.HypertextLinkClicked, AddressOf Me.TextControl_HypertextLinkClicked ' Opens the hyperlink when clicked.
            AddHandler Me.m_txTextControl.DocumentLinkClicked, AddressOf Me.TextControl_DocumentLinkClicked ' Scroll to the document target when clicked.
            AddHandler Loaded, AddressOf MainWindow_OpenHyperlink_Load ' Loads a linked file.
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' OpenHyperlink Method
        ' Open a new application instance if the hyperlink links a document. Open an internet browser if the hyperlink
        ' links to an http address.
        '
        ' Parameters:
        '		hyperlinkTarget:	The link to the http address or document that should be opened.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OpenHyperlink(ByVal hyperlinkTarget As String)
            If Not String.IsNullOrEmpty(hyperlinkTarget) Then
                Try
                    ' Create a Uri to determine the type of the linked target. 
                    Dim uriTarget As Uri
                    If Uri.TryCreate(hyperlinkTarget, UriKind.RelativeOrAbsolute, uriTarget) AndAlso uriTarget.IsAbsoluteUri OrElse Uri.TryCreate(Path.GetFullPath(hyperlinkTarget), UriKind.RelativeOrAbsolute, uriTarget) Then ' Handle relative file paths
                        ' Check whether the specified Uri is a file
                        If uriTarget.IsFile Then
                            ' Open the file by a type-corresponding application.
                            OpenFile(uriTarget)
                        Else
                            ' If it is not a file, the local system decides how to open the linked target. 
                            Process.Start(uriTarget.AbsoluteUri)
                        End If
                    End If
                    Return
                Catch
                End Try
            End If
            ' Inform the user if something went wrong.
            MessageBox.Show(Me, String.Format(My.Resources.MessageBox_OpenHyperlink_CouldNotOpenLink_Text, hyperlinkTarget), My.Resources.MessageBox_OpenHyperlink_CouldNotOpenLink_Caption, MessageBoxButton.OK, MessageBoxImage.Error)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' OpenFile Method
        ' Open the file that is linked by the specified uri.
        '
        ' Parameters:
        '		fileTarget: The uri that specifies the file to open.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OpenFile(ByVal fileTarget As Uri)
            Dim strFileToOpen = fileTarget.LocalPath

            ' Remove internal links inside the target document.
            Dim iPos = strFileToOpen.IndexOf("#")
            If iPos <> -1 Then
                strFileToOpen = strFileToOpen.Substring(0, iPos)
            End If

            ' Check whether the specified file exists.
            If Not File.Exists(strFileToOpen) Then
                MessageBox.Show(Me, String.Format(My.Resources.MessageBox_OpenHyperlink_FileDoesNotExist_Text, strFileToOpen), My.Resources.MessageBox_OpenHyperlink_FileDoesNotExist_Caption, MessageBoxButton.OK, MessageBoxImage.Error)
            Else
                ' If the file format is supported by TX Text Control...
                If IsSupportedDocumentFormat(strFileToOpen) Then
                    ' ... open the file with a new instance of this application.
                    OpenFileInNewInstance(strFileToOpen)
                Else
                    ' Otherwise open the file with the default application that 
                    ' is determined for the corresponding format.
                    Process.Start(strFileToOpen)
                End If
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' OpenFileInNewInstance
        ' Open the file by passed filepath in a new instance of this application.
        '
        ' Parameters:
        '      filePath:   The path to the file that should be opened in a new instance of this application
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OpenFileInNewInstance(ByVal filePath As String)
            ' Get running demo's exe path
            Dim strExePath As String = Assembly.GetEntryAssembly().Location

            ' Start new demo instance
            Dim pcPocess As Process = New Process()
            pcPocess.StartInfo.FileName = strExePath
            pcPocess.StartInfo.Arguments = """" & filePath & """"
            pcPocess.Start()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' IsSupportedDocumentFormat Method
        ' Checks whether the file format of the specified file is supported by TX Text Control.
        '
        ' Parameters:
        '		filePath:	The file to check.
        '
        ' Returns:	True if the specified file format is supported by TX Text Control. Otherwise false.
        '-----------------------------------------------------------------------------------------------------------
        Private Function IsSupportedDocumentFormat(ByVal filePath As String) As Boolean
            ' Check the extension of the file path.
            Select Case Path.GetExtension(filePath).ToLower()
                Case ".rtf", ".doc", ".docx", ".tx", ".xml", ".pdf", ".xlsx", ".txt"
                    Return True
            End Select
            Return False
        End Function
    End Class
End Namespace
