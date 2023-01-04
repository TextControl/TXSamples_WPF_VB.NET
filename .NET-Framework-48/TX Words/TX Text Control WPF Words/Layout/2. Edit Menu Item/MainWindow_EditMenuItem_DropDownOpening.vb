'-----------------------------------------------------------------------------------------------------------
' MainWindow_EditMenuItem_DropDownOpening.vb File
'
' Description: Provides all SubmenuOpened handlers associated with 'Edit' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' Edit_SubmenuOpened Handler
        '
        ' Updates the IsEnabled state and texts of 'Edit' drop down menu items.
        ' 
        ' Item: 'Edit'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Edit_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Get some current TextControl states
            Dim bCanCopy As Boolean = Me.m_txTextControl.CanCopy
            Dim bCanEdit As Boolean = Me.m_txTextControl.CanEdit
            Dim emEditMode As EditMode = Me.m_txTextControl.EditMode

            ' 'Undo'
            UpdateUndoText()
            Me.m_miEdit_Undo.IsEnabled = Me.m_txTextControl.CanUndo

            ' 'Redo'
            UpdateRedoText()
            Me.m_miEdit_Redo.IsEnabled = Me.m_txTextControl.CanRedo

            ' 'Cut'
            Me.m_miEdit_Cut.IsEnabled = bCanCopy AndAlso bCanEdit

            ' 'Copy'
            Me.m_miEdit_Copy.IsEnabled = bCanCopy

            ' 'Paste'
            Me.m_miEdit_Paste.IsEnabled = Me.m_txTextControl.CanPaste

            ' 'Select All'
            Me.m_miEdit_SelectAll.IsEnabled = bCanEdit OrElse emEditMode = EditMode.ReadAndSelect

            ' 'Replace...'
            Me.m_miEdit_Replace.IsEnabled = bCanEdit

            ' 'Protect Document
            Me.m_miEdit_ProtectDocument.IsChecked = emEditMode = EditMode.ReadAndSelect OrElse emEditMode = EditMode.ReadOnly

            ' 'Tracked Changes'
            Me.m_miEdit_TrackedChanges.IsEnabled = bCanEdit

            ' 'Comments'
            Me.m_miEdit_Comments.IsEnabled = bCanEdit
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Edit_Permissions_SubmenuOpened Handler
        '
        ' Updates the IsEnabled and checked state of 'Permissions' drop down menu items.
        ' 
        ' Item: 'Permissions'
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Edit_Permissions_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Get current document permissions
            Dim dpDocumentPermissions As DocumentPermissions = Me.m_txTextControl.DocumentPermissions

            ' Check the 'Permissions' drop down items

            ' Because formatting is not allowed in a read only document even if the corresponding document
            ' permissions are set, the corresponding items are unchecked when the DocumentPermissions.ReadOnly
            ' is set to true.
            Dim bIsReadOnly = dpDocumentPermissions.ReadOnly
            Me.m_miEdit_Permissions_AllowFormatting.IsChecked = Not bIsReadOnly AndAlso dpDocumentPermissions.AllowFormatting
            Me.m_miEdit_Permissions_AllowFormattingStyles.IsChecked = Not bIsReadOnly AndAlso dpDocumentPermissions.AllowFormattingStyles
            Me.m_miEdit_Permissions_AllowPrinting.IsChecked = dpDocumentPermissions.AllowPrinting
            Me.m_miEdit_Permissions_AllowCopy.IsChecked = dpDocumentPermissions.AllowCopy
            Me.m_miEdit_Permissions_AllowEditingFormFields.IsChecked = dpDocumentPermissions.AllowEditingFormFields
            Me.m_miEdit_Permissions_ReadOnly.IsChecked = bIsReadOnly

            ' Set the enable states of the 'Permissions' drop down items
            Dim bIsProtectedDocument As Boolean = Me.m_miEdit_ProtectDocument.IsChecked
            Me.m_miEdit_Permissions_AllowFormatting.IsEnabled = CSharpImpl.Assign(Me.m_miEdit_Permissions_AllowFormattingStyles.IsEnabled, Not bIsReadOnly AndAlso Not bIsProtectedDocument)
            Me.m_miEdit_Permissions_AllowPrinting.IsEnabled = CSharpImpl.Assign(Me.m_miEdit_Permissions_AllowCopy.IsEnabled, CSharpImpl.Assign(Me.m_miEdit_Permissions_AllowEditingFormFields.IsEnabled, CSharpImpl.Assign(Me.m_miEdit_Permissions_ReadOnly.IsEnabled, Not bIsProtectedDocument)))
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Edit_EditableRegions_Add_SubmenuOpened Handler
        '
        ' Updates the IsEnabled state of 'Add' drop down menu items.
        ' 
        ' Item: 'Add' of the 'Editable Regions' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Edit_EditableRegions_Add_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' Set the IsEnabled states of the 'Add' drop down items
                Dim colEditableRegions As EditableRegionCollection = Me.m_txTextControl.EditableRegions
                Dim rerCurrentEditableRegions As EditableRegion() = colEditableRegions.GetItems()
                ' The 'For [Current User]' item is IsEnabled if the current user is signed in and 
                ' no editable region was defined for the current user at the input position.
                Me.m_miEdit_EditableRegions_Add_ForCurrentUser.IsEnabled = m_uiCurrentUser IsNot Nothing AndAlso m_uiCurrentUser.IsSignedIn AndAlso (rerCurrentEditableRegions Is Nothing OrElse GetEditableRegion(m_strUserName, rerCurrentEditableRegions) Is Nothing)
                ' The 'For Everyone' item is IsEnabled if no corresponding editable region was
                ' at the current input position.
                Me.m_miEdit_EditableRegions_Add_ForEveryone.IsEnabled = rerCurrentEditableRegions Is Nothing OrElse GetEditableRegion("", rerCurrentEditableRegions) Is Nothing
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Edit_EditableRegions_Remove_SubmenuOpened Handler
        '
        ' Updates the IsEnabled state of 'Remove' drop down menu items.
        ' 
        ' Item: 'Remove' of the 'Editable Regions' drop down menu
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Edit_EditableRegions_Remove_SubmenuOpened(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If m_plTXLicense >= VersionInfo.ProductLevel.Professional Then
                ' Set the IsEnabled states of the 'Remove' drop down items
                Dim colEditableRegions As EditableRegionCollection = Me.m_txTextControl.EditableRegions
                Dim rerCurrentEditableRegions As EditableRegion() = colEditableRegions.GetItems()
                ' The 'For [Current User]' item is IsEnabled if the current user is signed in and 
                ' an editable region was defined for the current user at the input position.
                Me.m_miEdit_EditableRegions_Remove_ForCurrentUser.IsEnabled = m_uiCurrentUser IsNot Nothing AndAlso m_uiCurrentUser.IsSignedIn AndAlso rerCurrentEditableRegions IsNot Nothing AndAlso GetEditableRegion(m_strUserName, rerCurrentEditableRegions) IsNot Nothing
                ' The 'For Everyone' item is IsEnabled if a corresponding editable region was
                ' at the current input position.
                Me.m_miEdit_EditableRegions_Remove_ForEveryone.IsEnabled = rerCurrentEditableRegions IsNot Nothing AndAlso GetEditableRegion("", rerCurrentEditableRegions) IsNot Nothing
            End If
        End Sub
    End Class
End Namespace
