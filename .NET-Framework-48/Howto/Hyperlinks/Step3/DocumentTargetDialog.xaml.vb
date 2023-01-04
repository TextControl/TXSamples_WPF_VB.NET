'-----------------------------------------------------------------------------------------------------------
' DocumentTargetDialog.xaml.vb File
'
' Description:
'      A custom dialog to create or edit a document targets.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Public Class DocumentTargetDialog
    Inherits Window
    '-----------------------------------------------------------------------------------------------------------
    ' E N U M S
    '-----------------------------------------------------------------------------------------------------------
    Friend Enum DialogMode
        Insert
        Edit
    End Enum


    '-----------------------------------------------------------------------------------------------------------
    ' M E M B E R   V A R I A B L E S
    '-----------------------------------------------------------------------------------------------------------
    Private m_dtDocumentTarget As TXTextControl.DocumentTarget
    Private m_txTextControl As TXTextControl.WPF.TextControl
    Private m_dmDialogMode As DialogMode


    '-----------------------------------------------------------------------------------------------------------
    ' C O N S T R U C T O R
    '-----------------------------------------------------------------------------------------------------------
    Public Sub New(ByVal documentTarget As TXTextControl.DocumentTarget, ByVal textControl As TXTextControl.WPF.TextControl)
        m_dtDocumentTarget = documentTarget
        m_txTextControl = textControl
        Me.InitializeComponent()

        ' Determine the dialog mode.
        m_dmDialogMode = If(m_dtDocumentTarget Is Nothing, DialogMode.Insert, DialogMode.Edit)
    End Sub


    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' Window_Loaded Handler
    ' Updates the dialog layout according to the handled document target.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        ' Fill the 'Document targets at current input position' list box with all document targets that 
        ' are located at the current input position.
        Dim rdtCurrentDocumentTargets As TXTextControl.DocumentTarget() = m_txTextControl.DocumentTargets.GetItems()
        If rdtCurrentDocumentTargets IsNot Nothing Then
            Dim iCurrentIndex = -1
            For i = 0 To rdtCurrentDocumentTargets.Length - 1
                Dim rdtCurrentDocumentTarget = rdtCurrentDocumentTargets(i)
                Me.m_lbxCurrentDocumentTargets.Items.Add(New DocumentTargetItem(rdtCurrentDocumentTarget))
                ' Determine the index of that item that represents the document target that
                ' should be edited.
                If m_dmDialogMode = DialogMode.Edit AndAlso rdtCurrentDocumentTarget.GetHashCode() = m_dtDocumentTarget.GetHashCode() Then
                    iCurrentIndex = i
                End If
            Next
            ' Select the item that represents the document target that should be edited.
            Me.m_lbxCurrentDocumentTargets.SelectedIndex = iCurrentIndex
        End If

        ' Update the caption of the dialog, the visibility of the change button and the selection
        ' mode of the 'Document targets at current input position' combo box
        Select Case m_dmDialogMode
            Case DialogMode.Insert
                Title = "Insert Document Target"
                Me.m_btnChangeName.Visibility = Visibility.Collapsed
                Me.m_lbxCurrentDocumentTargets.IsEnabled = False
            Case DialogMode.Edit
                Title = "Edit Docoment Targets"
        End Select

        ' Fill the 'Document targets in document' list box with all document targets of the document.
        Dim colDocumentTargets = m_txTextControl.DocumentTargets
        For Each target As TXTextControl.DocumentTarget In colDocumentTargets
            Me.m_lbxAllDocumentTargets.Items.Add(New DocumentTargetItem(target))
        Next

        ' Update the enabled state of the 'OK' and 'Change Name' button.
        Me.m_btnOK.IsEnabled = CSharpImpl.Assign(Me.m_btnChangeName.IsEnabled, Me.m_tbxTargetName.Text.Trim().Length > 0)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' TargetName_TextChanged Handler
    ' Update the enabled state of the 'OK' and the 'Change Name' button when the text of the 'Target Name' 
    ' text box changed.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TargetName_TextChanged(sender As Object, e As TextChangedEventArgs)
        Me.m_btnOK.IsEnabled = CSharpImpl.Assign(Me.m_btnChangeName.IsEnabled, Me.m_tbxTargetName.Text.Trim().Length > 0)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' OK_Click Handler
    ' Create or edit a hyperlink when clicking the 'OK' button.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub OK_Click(sender As Object, e As RoutedEventArgs)
        Select Case m_dmDialogMode
            Case DialogMode.Insert
                ' The dialog is opened to create a document target:
                ' Create a new document target and insert it into the TextControl.
                Dim dtDocumentTarget As TXTextControl.DocumentTarget = New TXTextControl.DocumentTarget(Me.m_tbxTargetName.Text)
                dtDocumentTarget.Deleteable = Me.m_chbxCanBeDeleted.IsChecked.Value
                m_txTextControl.DocumentTargets.Add(dtDocumentTarget)
            Case DialogMode.Edit
                ' The dialog is opened to edit a document target:
                ' Update the TargetName and the Deleteable property values of the document target.
                For Each item As DocumentTargetItem In Me.m_lbxCurrentDocumentTargets.Items
                    item.DocumentTarget.TargetName = item.DisplayedText
                    item.DocumentTarget.Deleteable = item.IsDocumentTargetDeletable
                Next
        End Select
        ' Close the dialog.
        DialogResult = True
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' CurrentDocumentTargets_ItemSelected Handler
    ' Update the text of the 'Target Name' text box with the displayed text of the new selected item. Furthermore,
    ' the 'Can be deleted during editing' check box is updated.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub CurrentDocumentTargets_ItemSelected(sender As Object, e As SelectionChangedEventArgs)
        If Me.m_lbxCurrentDocumentTargets.SelectedIndex <> -1 Then
            Dim dtiSelectedItem As DocumentTargetItem = TryCast(Me.m_lbxCurrentDocumentTargets.SelectedItem, DocumentTargetItem)
            Me.m_tbxTargetName.Text = dtiSelectedItem.DisplayedText
            Me.m_chbxCanBeDeleted.IsChecked = dtiSelectedItem.IsDocumentTargetDeletable
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' ChangeName_Click Handler
    ' Change the DisplayedText property value of the selected item inside the 'Document targets at current input 
    ' position' list box to the text of the 'Target Name' text box when the 'Change Name' button is clicked.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ChangeName_Click(sender As Object, e As RoutedEventArgs)
        Dim dtiNewItem As DocumentTargetItem = New DocumentTargetItem(TryCast(Me.m_lbxCurrentDocumentTargets.SelectedItem, DocumentTargetItem).DocumentTarget)
        dtiNewItem.DisplayedText = Me.m_tbxTargetName.Text
        Me.m_lbxCurrentDocumentTargets.Items(Me.m_lbxCurrentDocumentTargets.SelectedIndex) = dtiNewItem
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' CanBeDeleted_CheckedChanged Handler
    ' Change the IsDocumentTargetDeletable property value of the selected item inside the 'Document targets at  
    ' current input position' list box to the check state of the 'Can be deleted during editing' check box.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub CanBeDeleted_CheckedChanged(sender As Object, e As RoutedEventArgs)
        If Me.m_lbxCurrentDocumentTargets.SelectedIndex <> -1 Then
            Dim dtiSelectedItem As DocumentTargetItem = TryCast(Me.m_lbxCurrentDocumentTargets.SelectedItem, DocumentTargetItem)
            dtiSelectedItem.IsDocumentTargetDeletable = Me.m_chbxCanBeDeleted.IsChecked.Value
        End If
    End Sub


    '-----------------------------------------------------------------------------------------------------------
    ' S U P P O R T I N G   C L A S S E S
    '-----------------------------------------------------------------------------------------------------------
    Private Class CSharpImpl
        Shared Function Assign(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class
End Class
