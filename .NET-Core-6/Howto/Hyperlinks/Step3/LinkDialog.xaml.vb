'-----------------------------------------------------------------------------------------------------------
' LinkDialog.xaml.vb File
'
' Description:
'      A custom dialog to create or edit a link.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Public Class LinkDialog
    '-----------------------------------------------------------------------------------------------------------
    ' E N U M S
    '-----------------------------------------------------------------------------------------------------------
    Friend Enum DialogMode
        InsertLink
        EditHyperlink
        EditDocumentLink
    End Enum


    '-----------------------------------------------------------------------------------------------------------
    ' M E M B E R   V A R I A B L E S
    '-----------------------------------------------------------------------------------------------------------
    Private m_dmDialogMode As DialogMode
    Private m_tfLink As TXTextControl.TextField
    Private m_txTextControl As TXTextControl.WPF.TextControl


    '-----------------------------------------------------------------------------------------------------------
    ' C O N S T R U C T O R
    '-----------------------------------------------------------------------------------------------------------
    Public Sub New(ByVal link As TXTextControl.TextField, ByVal textControl As TXTextControl.WPF.TextControl)
        m_tfLink = link
        m_txTextControl = textControl
        Me.InitializeComponent()

        ' Determine the dialog mode.
        m_dmDialogMode = If(m_tfLink Is Nothing, DialogMode.InsertLink, If(TypeOf m_tfLink Is TXTextControl.HypertextLink, DialogMode.EditHyperlink, DialogMode.EditDocumentLink))
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' Window_Loaded Handler
    ' Updates the dialog layout according to the handled link.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        ' Update the dialog layout considering the set dialog mode.
        Select Case m_dmDialogMode
            Case DialogMode.InsertLink
                Me.m_cmbxDocumentTargets.Visibility = Visibility.Collapsed
                Title = "Insert Link"
            Case DialogMode.EditHyperlink
                Title = "Edit Hyperlink"
                Me.m_grdLinkType.Visibility = Visibility.Collapsed
                Me.m_cmbxDocumentTargets.Visibility = Visibility.Collapsed
            Case DialogMode.EditDocumentLink
                Title = "Edit Document Link"
                Me.m_grdLinkType.Visibility = Visibility.Collapsed
                Me.m_tbxHyperlink.Visibility = Visibility.Collapsed
        End Select

        ' Fill the document targets combo box with references to the document targets
        ' of the document.
        If m_dmDialogMode <> DialogMode.EditHyperlink Then
            Dim colDocumentTargets = m_txTextControl.DocumentTargets
            For Each target As TXTextControl.DocumentTarget In colDocumentTargets
                Me.m_cmbxDocumentTargets.Items.Add(New DocumentTargetItem(target))
            Next
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Type_CheckedChanged Handler
    ' Update the visibility of the corresponding control when the type radio button checked state changed.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Type_CheckedChanged(sender As Object, e As RoutedEventArgs)
        If IsLoaded Then
            If Not Me.m_rbtnTypeHyperlink.IsChecked.Value Then
                ' The document targets combo box is displayed when the 'Document Link' radio button is toggled.
                Me.m_tbxHyperlink.Visibility = Visibility.Collapsed
                Me.m_cmbxDocumentTargets.Visibility = Visibility.Visible
            Else
                ' The text box is displayed to enter a hyperlink when the 'Hyperlink' radio button is toggled.
                Me.m_cmbxDocumentTargets.Visibility = Visibility.Collapsed
                Me.m_tbxHyperlink.Visibility = Visibility.Visible
            End If

            ' Update the enabled state of the OK button.
            Me.m_btnOK.IsEnabled = IsValidLink()
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' LinkParameter_Changed Handler
    ' Update the enabled state of the 'OK' button when the text of any text boxes changed.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub LinkParameter_Changed(sender As Object, e As TextChangedEventArgs)
        Me.m_btnOK.IsEnabled = IsValidLink()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' LinkParameter_Changed Handler
    ' Update the enabled state of the 'OK' button when the selected item of the document targets combo box changed.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub DocumentTargets_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Me.m_btnOK.IsEnabled = IsValidLink()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' OK_Click Handler
    ' Create or edit a link when clicking the 'OK' button.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub OK_Click(sender As Object, e As RoutedEventArgs)
        ' Consider the current mode.
        Select Case m_dmDialogMode
            Case DialogMode.InsertLink
                ' The dialog is opened to create a link:
                If Me.m_rbtnTypeHyperlink.IsChecked.Value Then
                    ' Create a new hyperlink and insert it into the TextControl.
                    Dim hlHypertextLink As TXTextControl.HypertextLink = New TXTextControl.HypertextLink(Me.m_tbxDisplayedText.Text, Me.m_tbxHyperlink.Text)
                    hlHypertextLink.DoubledInputPosition = True
                    m_txTextControl.HypertextLinks.Add(hlHypertextLink)
                Else
                    ' Create a new document link and insert it into the TextControl.
                    Dim dlDocumentLink As TXTextControl.DocumentLink = New TXTextControl.DocumentLink(Me.m_tbxDisplayedText.Text, TryCast(Me.m_cmbxDocumentTargets.SelectedItem, DocumentTargetItem).DocumentTarget)
                    dlDocumentLink.DoubledInputPosition = True
                    m_txTextControl.DocumentLinks.Add(dlDocumentLink)
                End If
            Case DialogMode.EditHyperlink
                ' The dialog is opened to edit a hyperlink:
                ' Update the text of the hyperlink.
                Dim hlHypertextLink As TXTextControl.HypertextLink = TryCast(m_tfLink, TXTextControl.HypertextLink)
                hlHypertextLink.Text = Me.m_tbxDisplayedText.Text
                hlHypertextLink.Target = Me.m_tbxHyperlink.Text
            Case DialogMode.EditDocumentLink
                ' The dialog is opened to edit a document link:
                ' Update the text and the document target of the document link.
                Dim dlDocumentLink As TXTextControl.DocumentLink = TryCast(m_tfLink, TXTextControl.DocumentLink)
                dlDocumentLink.Text = Me.m_tbxDisplayedText.Text
                dlDocumentLink.DocumentTarget = TryCast(Me.m_cmbxDocumentTargets.SelectedItem, DocumentTargetItem).DocumentTarget
        End Select

        ' Close the dialog.
        DialogResult = True
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' M E T H O D S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' IsValidLink Method
    ' Returns a value indicating whether both the 'Displayed Text' text box contains text and the link specific
    ' control (hyperlink text box or document link combo box) contains a valid value.
    '
    ' Returns:		True, if both the 'Displayed Text' text box contains text and the link specific control 
    '				(hyperlink text box or document link combo box) contains a valid value.
    '				Otherwise false.
    '-----------------------------------------------------------------------------------------------------------
    Private Function IsValidLink() As Boolean
        Return Me.m_tbxDisplayedText.Text.Trim().Length > 0 AndAlso (If(Me.m_cmbxDocumentTargets.Visibility = Visibility.Visible, Me.m_cmbxDocumentTargets.SelectedIndex <> -1, Me.m_tbxHyperlink.Text.Trim().Length > 0))
    End Function
End Class
