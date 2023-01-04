'-----------------------------------------------------------------------------------------------------------
' DeleteAndGoToTargetDialog.xaml.vb File
'
' Description:
'      A custom dialog to delete or navigate to document targets.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Public Class DeleteAndGoToTargetDialog
    '-----------------------------------------------------------------------------------------------------------
    ' M E M B E R   V A R I A B L E S
    '-----------------------------------------------------------------------------------------------------------
    Private m_txTextControl As TXTextControl.WPF.TextControl
    Private m_lstTagetsToDelete As List(Of TXTextControl.DocumentTarget) = New List(Of TXTextControl.DocumentTarget)()


    '-----------------------------------------------------------------------------------------------------------
    ' C O N S T R U C T O R
    '-----------------------------------------------------------------------------------------------------------
    Public Sub New(ByVal textControl As TXTextControl.WPF.TextControl)
        m_txTextControl = textControl
        Me.InitializeComponent()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' Window_Loaded Handler
    ' Updates the dialog layout.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        ' Fill the 'Document Targets' list box with all document targets of the document.
        Dim colDocumentTargets = m_txTextControl.DocumentTargets
        For Each documentTarget As TXTextControl.DocumentTarget In colDocumentTargets
            Me.m_lbxDocumentTargets.Items.Add(New DocumentTargetItem(documentTarget))
        Next

        ' Select the first item.
        If Me.m_lbxDocumentTargets.Items.Count > 0 Then
            Me.m_lbxDocumentTargets.SelectedIndex = 0
        End If
    End Sub


    '-----------------------------------------------------------------------------------------------------------
    ' DocumentTargets_SelectedIndexChanged Handler
    ' Update the enabled state of the 'Delete' and 'Go To' buttons when the selected index of the 
    ' 'Document Targets' list box changed.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub DocumentTargets_SelectedIndexChanged(sender As Object, e As SelectionChangedEventArgs)
        Me.m_btnDelete.IsEnabled = CSharpImpl.Assign(Me.m_btnGoTo.IsEnabled, Me.m_lbxDocumentTargets.SelectedIndex <> -1)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Delete_Click Handler
    ' Remove the selected item from the 'Document Targets' list box.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Delete_Click(sender As Object, e As RoutedEventArgs)
        ' Remember the current index.
        Dim iSelectedIndex As Integer = Me.m_lbxDocumentTargets.SelectedIndex

        ' Get the current selected item.
        Dim dtiItemToDelete As DocumentTargetItem = TryCast(Me.m_lbxDocumentTargets.SelectedItem, DocumentTargetItem)
        m_lstTagetsToDelete.Add(dtiItemToDelete.DocumentTarget) ' Remember that item.
        Me.m_lbxDocumentTargets.Items.Remove(dtiItemToDelete) ' Remove that item from the 'Document Targets' list box.

        ' Select a new item.
        If Me.m_lbxDocumentTargets.Items.Count > 0 Then
            Me.m_lbxDocumentTargets.SelectedIndex = Math.Max(0, Math.Min(iSelectedIndex, Me.m_lbxDocumentTargets.Items.Count - 1))
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' GoTo_Click Handler
    ' Scroll to the document target that is represented by the selected item inside the 
    ' 'Document Targets' list box.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub GoTo_Click(sender As Object, e As RoutedEventArgs)
        TryCast(Me.m_lbxDocumentTargets.SelectedItem, DocumentTargetItem).DocumentTarget.ScrollTo()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' OK_Click Handler
    ' Delete all corresponding document targets that were removed from the 'Document Targets' list box.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub OK_Click(sender As Object, e As RoutedEventArgs)
        If m_lstTagetsToDelete.Count > 0 Then
            Dim colDocumentTargets = m_txTextControl.DocumentTargets
            For Each documentTarget In m_lstTagetsToDelete
                colDocumentTargets.Remove(documentTarget)
            Next
        End If

        ' Close the dialog.
        DialogResult = True
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
