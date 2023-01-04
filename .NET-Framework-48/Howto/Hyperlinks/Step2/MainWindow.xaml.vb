'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'      Sample project that is related to the 'Howto: Use Hypertext Links -> Step 2: Adding a Dialog Box for 
'	   Inserting Hypertext Links ' article inside the 'Windows Presentation Foundation User's Guide'.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Class MainWindow
    Inherits Window
    '-----------------------------------------------------------------------------------------------------------
    ' M E M B E R   V A R I A B L E S
    '-----------------------------------------------------------------------------------------------------------
    Private m_hlCurrentHyperlink As TXTextControl.HypertextLink = Nothing

    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' Hyperlinks_DropDownOpening Handler
    ' When opening the 'Hyperlinks' menu item drop down, check whether there is a hyperlink at the current 
    ' input position and update the enabled state of the 'Insert...' and 'Edit...' items.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Hyperlinks_DropDownOpening(sender As Object, e As RoutedEventArgs)
        ' Get the hyperlink at the current input position.
        m_hlCurrentHyperlink = Me.m_txTextControl.HypertextLinks.GetItem()
        ' Disable the 'Insert...' item if there is a hyperlink at the
        ' current input position.
        Me.m_tmiInsertHyperlink.IsEnabled = m_hlCurrentHyperlink Is Nothing
        ' Disable the 'Edit...' item if there is no hyperlink at the
        ' current input position.
        Me.m_tmiEditHyperlink.IsEnabled = m_hlCurrentHyperlink IsNot Nothing
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' OpenHyperlinkDialog_Click Handler
    ' Open the custom hyperlink dialog to create or edit a hyperlink.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub OpenHyperlinkDialog_Click(sender As Object, e As RoutedEventArgs)
        ' Open the custom hyperlink dialog to create or edit a hyperlink.
        Dim dlgInsertHyperlinkDialog As CustomHyperlinkDialog = New CustomHyperlinkDialog(m_hlCurrentHyperlink, Me.m_txTextControl)
        dlgInsertHyperlinkDialog.Owner = Me
        dlgInsertHyperlinkDialog.ShowDialog()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_TextFieldCreated Handler
    ' Update the underline style and fore color settings when a hyperlink is created.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_TextFieldCreated(sender As Object, e As TXTextControl.TextFieldEventArgs)
        ' Update the underline style and fore color settings of the created hyperlink.
        If TypeOf e.TextField Is TXTextControl.HypertextLink Then
            Me.HighlightHyperlink(TryCast(e.TextField, TXTextControl.HypertextLink), Me.m_tmiShowHyperlinks.IsChecked)
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' ShowHyperlinks_Click Handler
    ' Step through all hyperlinks to update the underline style and fore color settings.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ShowHyperlinks_Click(sender As Object, e As RoutedEventArgs)
        Me.m_txTextControl.BeginUndoAction("Show Hyperlinks")
        ' Step through all hyperlinks to update the underline style and fore color settings.
        Dim colHyperlinks As TXTextControl.HypertextLinkCollection = Me.m_txTextControl.HypertextLinks
        For Each link As TXTextControl.HypertextLink In colHyperlinks
            Me.HighlightHyperlink(link, Me.m_tmiShowHyperlinks.IsChecked)
        Next
        Me.m_txTextControl.EndUndoAction()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
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


    '-----------------------------------------------------------------------------------------------------------
    ' M E T H O D S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' HighlightHyperlink Method
    ' Determines for the specified hyperlink whether or not the text should be
    ' formatted with blue fore color and a single underline style.
    '
    ' Parameters:
    '		link		The hyperlink to handle.
    '		highlight	A value that indicates whether or not the hyperlink 
    '					should be highlighted.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub HighlightHyperlink(ByVal link As TXTextControl.HypertextLink, ByVal highlight As Boolean)
        ' Create a selection object to determine the color and underline style
        ' of the hyperlink text.
        Dim selLink As TXTextControl.Selection = New TXTextControl.Selection(link.Start - 1, link.Length)

        ' Determine the fore color and underline style settings.
        If highlight Then
            selLink.ForeColor = System.Drawing.Color.Blue
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
