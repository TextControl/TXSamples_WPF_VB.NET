' -------------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'      Sample project that is related to the 'Howto: Use Hypertext Links -> Step 3: Adding Document Targets' 
'	   article inside the 'Windows Presentation Foundation User's Guide'.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Class MainWindow
    ' -------------------------------------------------------------------------------------------------------------
    ' M E M B E R   V A R I A B L E S
    '-----------------------------------------------------------------------------------------------------------
    Private m_tfCurrentLink As TXTextControl.TextField = Nothing
    Private m_dtCurrentDocumentTarget As TXTextControl.DocumentTarget = Nothing


    ' -------------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    ' -------------------------------------------------------------------------------------------------------------
    ' 'File' Drop Down
    '-----------------------------------------------------------------------------------------------------------

    ' -------------------------------------------------------------------------------------------------------------
    ' New_Click Handler
    ' Deletes the entire contents of a Text Control.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub New_Click(sender As Object, e As RoutedEventArgs)
        Me.m_txTextControl.ResetContents()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' Open_Click Handler
    ' Opens a built-in dialog to load a document into the TextControl.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Open_Click(sender As Object, e As RoutedEventArgs)
        Me.m_txTextControl.Load()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' SaveAs_Click Handler
    ' Opens a built-in dialog to save the document.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub SaveAs_Click(sender As Object, e As RoutedEventArgs)
        Me.m_txTextControl.Save()
    End Sub


    ' -------------------------------------------------------------------------------------------------------------
    ' 'Links' Drop Down
    '-----------------------------------------------------------------------------------------------------------

    ' -------------------------------------------------------------------------------------------------------------
    ' Links_DropDownOpening Handler
    ' When opening the 'Links' menu item drop down, check whether there is a text field (base class of 
    ' HypertextLink and DocumentLink) at the current input position and update the enabled state of the 
    ' corresponding 'Insert...' and 'Edit...' items.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Links_DropDownOpening(sender As Object, e As RoutedEventArgs)
        ' Get the text field at the current input position.
        m_tfCurrentLink = Me.m_txTextControl.TextFields.GetItem()
        ' Disable the 'Insert...' item if there is a text field at the
        ' current input position.
        Me.m_tmiInsertLink.IsEnabled = m_tfCurrentLink Is Nothing
        ' Disable the 'Edit...' item if there is no hyperlink or 
        ' document linkat the current input position.
        Me.m_tmiEditLink.IsEnabled = m_tfCurrentLink IsNot Nothing AndAlso (TypeOf m_tfCurrentLink Is TXTextControl.HypertextLink OrElse TypeOf m_tfCurrentLink Is TXTextControl.DocumentLink)
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' OpenCustomLinkDialog_Click Handler
    ' Open the custom link dialog to create or edit a link.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub OpenCustomLinkDialog_Click(sender As Object, e As RoutedEventArgs)
        ' Open the custom link dialog to create or edit a link.
        Dim dlgLinkDialog As LinkDialog = New LinkDialog(m_tfCurrentLink, Me.m_txTextControl)
        dlgLinkDialog.Owner = Me
        dlgLinkDialog.ShowDialog()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' TextControl_TextFieldCreated Handler
    ' Update the underline style and fore color settings when a hyperlink or document link is created.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_TextFieldCreated(sender As Object, e As TXTextControl.TextFieldEventArgs)
        ' Update the underline style and fore color settings of the created link.
        If TypeOf e.TextField Is TXTextControl.HypertextLink OrElse TypeOf e.TextField Is TXTextControl.DocumentLink Then
            Me.HighlightLink(e.TextField, Me.m_tmiShowHyperlinks.IsChecked)
        End If
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' ShowHyperlinks_Click Handler
    ' Step through all hyperlinks to update the underline style and fore color settings.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ShowHyperlinks_Click(sender As Object, e As RoutedEventArgs)
        Me.m_txTextControl.BeginUndoAction("Show Hyperlinks")
        ' Step through all hyperlinks to update the underline style and fore color settings.
        Dim colHyperlinks As TXTextControl.HypertextLinkCollection = Me.m_txTextControl.HypertextLinks
        For Each link As TXTextControl.HypertextLink In colHyperlinks
            Me.HighlightLink(link, Me.m_tmiShowHyperlinks.IsChecked)
        Next
        Me.m_txTextControl.EndUndoAction()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' ShowDocumentLinks_Click Handler
    ' Step through all document links to update the underline style and fore color settings.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ShowDocumentLinks_Click(sender As Object, e As RoutedEventArgs)
        Me.m_txTextControl.BeginUndoAction("Show Document Links")
        ' Step through all document links to update the underline style and fore color settings.
        Dim colDocumentLinks As TXTextControl.DocumentLinkCollection = Me.m_txTextControl.DocumentLinks
        For Each link As TXTextControl.DocumentLink In colDocumentLinks
            Me.HighlightLink(link, Me.m_tmiShowDocumentLinks.IsChecked)
        Next
        Me.m_txTextControl.EndUndoAction()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' TextControl_HypertextLinkClicked Handler
    ' Open the clicked link by the system's default application to open a hyperlink.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_HypertextLinkClicked(sender As Object, e As TXTextControl.HypertextLinkEventArgs)
        ' Get the hyperlink target.
        Dim strHyperlink = e.HypertextLink.Target
        ' Update target to a correct hyperlink if necessary.
        If Not strHyperlink.StartsWith("http://") Then
            strHyperlink = "http://" & strHyperlink
        End If
        ' Create an URI object and try to open the link by the system's default application
        ' to open a hyperlink.
        Dim uriTarget As Uri
        If Uri.TryCreate(strHyperlink, UriKind.RelativeOrAbsolute, uriTarget) AndAlso uriTarget.IsAbsoluteUri Then
            Process.Start(uriTarget.AbsoluteUri)
        End If
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' TextControl_DocumentLinkClicked Handler
    ' Scroll to the target of the clicked document link.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_DocumentLinkClicked(sender As Object, e As TXTextControl.DocumentLinkEventArgs)
        Dim dtTarget = e.DocumentLink.DocumentTarget
        If dtTarget IsNot Nothing Then
            dtTarget.ScrollTo()
        End If
    End Sub


    ' -------------------------------------------------------------------------------------------------------------
    ' 'Document Targets' Drop Down
    '-----------------------------------------------------------------------------------------------------------

    ' -------------------------------------------------------------------------------------------------------------
    ' DocumentTargets_DropDownOpening Handler
    ' When opening the 'Document Targets' menu item drop down, check whether there is a document target at 
    ' the current input position and update the enabled state of the corresponding 'Insert..' and 'Edit...' items.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub DocumentTargets_DropDownOpening(sender As Object, e As RoutedEventArgs)
        m_dtCurrentDocumentTarget = Me.m_txTextControl.DocumentTargets.GetItem()
        Me.m_tmiEditDocumentTarget.IsEnabled = m_dtCurrentDocumentTarget IsNot Nothing
        Me.m_tmiShowDocumentTargets.IsChecked = Me.m_txTextControl.DocumentTargetMarkers
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' InsertDocumentTarget_Click Handler
    ' Open the document target dialog to create a document target.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub InsertDocumentTarget_Click(sender As Object, e As RoutedEventArgs)
        Dim dlgDocumentTargetDialog As DocumentTargetDialog = New DocumentTargetDialog(Nothing, Me.m_txTextControl)
        dlgDocumentTargetDialog.Owner = Me
        dlgDocumentTargetDialog.ShowDialog()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' EditDocumentTarget_Click Handler
    ' Open the document target dialog to edit a document target.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub EditDocumentTarget_Click(sender As Object, e As RoutedEventArgs)
        Dim dlgDocumentTargetDialog As DocumentTargetDialog = New DocumentTargetDialog(m_dtCurrentDocumentTarget, Me.m_txTextControl)
        dlgDocumentTargetDialog.Owner = Me
        dlgDocumentTargetDialog.ShowDialog()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' DeleteAndGoToTarget_Click Handler
    ' Open the delete and go to target dialog to delete or navigate to a document target.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub DeleteAndGoToTarget_Click(sender As Object, e As RoutedEventArgs)
        Dim dlgDeleteAndGoToTargetDialog As DeleteAndGoToTargetDialog = New DeleteAndGoToTargetDialog(Me.m_txTextControl)
        dlgDeleteAndGoToTargetDialog.Owner = Me
        dlgDeleteAndGoToTargetDialog.ShowDialog()
    End Sub

    ' -------------------------------------------------------------------------------------------------------------
    ' ShowDocumentTargets_Click Handler
    ' Sets a value indicating that markers for document targets are shown or not.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ShowDocumentTargets_Click(sender As Object, e As RoutedEventArgs)
        Me.m_txTextControl.DocumentTargetMarkers = Me.m_tmiShowDocumentTargets.IsChecked
    End Sub


    ' -------------------------------------------------------------------------------------------------------------
    ' M E T H O D S
    '-----------------------------------------------------------------------------------------------------------

    ' -------------------------------------------------------------------------------------------------------------
    ' HighlightLink Method
    ' Determines for the specified link whether or not the text should be formatted with
    ' blue (hyperlink) or green (document link) fore color and a single underline style.
    '
    ' Parameters:
    '		link		The link to handle.
    '		highlight	A value that indicates whether or not the link 
    '					should be highlighted.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub HighlightLink(ByVal link As TXTextControl.TextField, ByVal highlight As Boolean)
        ' Create a selection object to determine the color and underline style
        ' of the link text.
        Dim selLink As TXTextControl.Selection = New TXTextControl.Selection(link.Start - 1, link.Length)

        ' Determine the fore color and underline style settings.
        If highlight Then
            selLink.ForeColor = If(TypeOf link Is TXTextControl.HypertextLink, System.Drawing.Color.Blue, System.Drawing.Color.Green)
            selLink.Underline = TXTextControl.FontUnderlineStyle.Single
        Else
            selLink.ForeColor = System.Drawing.Color.Black
            selLink.Underline = TXTextControl.FontUnderlineStyle.None
        End If

        ' Apply the settings by adopting the selection.
        Me.m_txTextControl.Selection = selLink

        ' Set the input position inside the hyperlink to prevent adopting the set formatting when
        ' typing right behind the hyperlink.
        Me.m_txTextControl.InputPosition = New TXTextControl.InputPosition(selLink.Start + selLink.Length, TXTextControl.TextFieldPosition.InsideTextField)
    End Sub
End Class


' -------------------------------------------------------------------------------------------------------------
' S U P P O R T I N G   C L A S S E S
'-----------------------------------------------------------------------------------------------------------

' -------------------------------------------------------------------------------------------------------------
' DocumentTargetItem class
' Used as item inside a combo or list box to display an obejct of type DocumentTarget. 
'-----------------------------------------------------------------------------------------------------------
Friend Class DocumentTargetItem
    ' Member Variables
    Private m_dtDocumentTarget As TXTextControl.DocumentTarget
    Private m_strDisplayedText As String
    Private m_bIsDocumentTargetDeletable As Boolean

    ' Constructor
    Friend Sub New(ByVal documentTarget As TXTextControl.DocumentTarget)
        m_dtDocumentTarget = documentTarget
        m_strDisplayedText = m_dtDocumentTarget.TargetName
        m_bIsDocumentTargetDeletable = documentTarget.Deleteable
    End Sub

    ' Properties
    Friend Property DisplayedText As String
        Get
            Return m_strDisplayedText
        End Get
        Set(ByVal value As String)
            m_strDisplayedText = value
        End Set
    End Property

    Friend ReadOnly Property DocumentTarget As TXTextControl.DocumentTarget
        Get
            Return m_dtDocumentTarget
        End Get
    End Property

    Friend Property IsDocumentTargetDeletable As Boolean
        Get
            Return m_bIsDocumentTargetDeletable
        End Get
        Set(ByVal value As Boolean)
            m_bIsDocumentTargetDeletable = value
        End Set
    End Property

    ' Overridden Methods
    Public Overrides Function ToString() As String
        Return m_strDisplayedText
    End Function
End Class