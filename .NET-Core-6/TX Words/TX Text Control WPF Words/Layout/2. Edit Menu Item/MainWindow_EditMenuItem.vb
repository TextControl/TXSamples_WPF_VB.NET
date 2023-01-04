'-----------------------------------------------------------------------------------------------------------
' MainWindow_EditMenuItem.vb File
'
' Description: Provides methods to set the layout of the 'Edit' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' SetEditItemsTexts Method
        '
        ' Sets the texts of the 'Edit' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetEditItemsTexts()
            ' 'Edit'
            m_miEdit.Header = My.Resources.Item_Edit_Text

            ' 'Undo'
            m_miEdit_Undo.Header = My.Resources.Item_Edit_Undo_Text

            ' 'Redo'
            m_miEdit_Redo.Header = My.Resources.Item_Edit_Redo_Text

            ' 'Cut'
            m_miEdit_Cut.Header = My.Resources.Item_Edit_Cut_Text

            ' 'Copy'
            m_miEdit_Copy.Header = My.Resources.Item_Edit_Copy_Text

            ' 'Paste'
            m_miEdit_Paste.Header = My.Resources.Item_Edit_Paste_Text

            ' 'Select All'
            m_miEdit_SelectAll.Header = My.Resources.Item_Edit_SelectAll_Text

            ' 'Find...'
            m_miEdit_Find.Header = My.Resources.Item_Edit_Find_Text

            ' 'Replace...'
            m_miEdit_Replace.Header = My.Resources.Item_Edit_Replace_Text

            ' 'Permissions'
            m_miEdit_Permissions.Header = My.Resources.Item_Edit_Permissions_Text
            m_miEdit_Permissions_AllowFormatting.Header = My.Resources.Item_Edit_Permissions_AllowFormatting_Text
            m_miEdit_Permissions_AllowFormattingStyles.Header = My.Resources.Item_Edit_Permissions_AllowFormattingStyles_Text
            m_miEdit_Permissions_AllowPrinting.Header = My.Resources.Item_Edit_Permissions_AllowPrinting_Text
            m_miEdit_Permissions_AllowCopy.Header = My.Resources.Item_Edit_Permissions_AllowCopy_Text
            m_miEdit_Permissions_AllowEditingFormFields.Header = My.Resources.Item_Edit_Permissions_AllowEditingFormFields_Text
            m_miEdit_Permissions_ReadOnly.Header = My.Resources.Item_Edit_Permissions_ReadOnly_Text

            ' 'Editable Regions'
            m_miEdit_EditableRegions.Header = My.Resources.Item_Edit_EditableRegions_Text
            m_miEdit_EditableRegions_Add.Header = My.Resources.Item_Edit_EditableRegions_Add_Text
            Me.SetItemText(Me.m_miEdit_EditableRegions_Add_ForCurrentUser, m_strUserName)
            m_miEdit_EditableRegions_Add_ForEveryone.Header = My.Resources.Item_Edit_EditableRegions_Add_ForEveryone_Text
            m_miEdit_EditableRegions_Remove.Header = My.Resources.Item_Edit_EditableRegions_Remove_Text
            Me.SetItemText(Me.m_miEdit_EditableRegions_Remove_ForCurrentUser, m_strUserName)
            m_miEdit_EditableRegions_Remove_ForEveryone.Header = My.Resources.Item_Edit_EditableRegions_Remove_ForEveryone_Text

            ' 'Protect Document'
            m_miEdit_ProtectDocument.Header = My.Resources.Item_Edit_ProtectDocument_Text

            ' 'Tracked Changes'
            m_miEdit_TrackedChanges.Header = My.Resources.Item_Edit_TrackedChanges_Text
            m_miEdit_TrackedChanges_TrackChanges.Header = My.Resources.Item_Edit_TrackedChanges_TrackChanges_Text
            m_miEdit_TrackedChanges_ReviewTrackedChanges.Header = My.Resources.Item_Edit_TrackedChanges_ReviewTrackedChanges_Text

            ' 'Comments'
            m_miEdit_Comments.Header = My.Resources.Item_Edit_Comments_Text
            m_miEdit_Comments_AddComment.Header = My.Resources.Item_Edit_Comments_AddComment_Text
            m_miEdit_Comments_ReviewComments.Header = My.Resources.Item_Edit_Comments_ReviewComments_Text
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SetEditItemsImages Method
        '
        ' Sets the images of the 'Edit' menu items.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetEditItemsImages()
            ' 'Undo'
            Me.SetItemImage(Me.m_miEdit_Undo, ResourceProvider.GeneralItem.TXITEM_Undo.ToString())

            ' 'Redo'
            Me.SetItemImage(Me.m_miEdit_Redo, ResourceProvider.GeneralItem.TXITEM_Redo.ToString())

            ' 'Cut'
            Me.SetItemImage(Me.m_miEdit_Cut, RibbonFormattingTab.RibbonItem.TXITEM_Cut.ToString())

            ' 'Copy'
            Me.SetItemImage(Me.m_miEdit_Copy, RibbonFormattingTab.RibbonItem.TXITEM_Copy.ToString())

            ' 'Paste'
            Me.SetItemImage(Me.m_miEdit_Paste, RibbonFormattingTab.RibbonItem.TXITEM_Paste.ToString())

            ' 'Select All'
            Me.SetItemImage(Me.m_miEdit_SelectAll, RibbonFormattingTab.RibbonItem.TXITEM_SelectAll.ToString())

            ' 'Find...'
            Me.SetItemImage(Me.m_miEdit_Find, RibbonFormattingTab.RibbonItem.TXITEM_Find.ToString())

            ' 'Replace...'
            Me.SetItemImage(Me.m_miEdit_Replace, RibbonFormattingTab.RibbonItem.TXITEM_Replace.ToString())

            ' 'Permissions'
            Me.SetItemImage(Me.m_miEdit_Permissions, RibbonPermissionsTab.RibbonItem.TXITEM_ReadOnly.ToString())
            Me.SetItemImage(Me.m_miEdit_Permissions_AllowFormatting, RibbonPermissionsTab.RibbonItem.TXITEM_AllowFormatting.ToString())
            Me.SetItemImage(Me.m_miEdit_Permissions_AllowFormattingStyles, RibbonPermissionsTab.RibbonItem.TXITEM_AllowFormattingStyles.ToString())
            Me.SetItemImage(Me.m_miEdit_Permissions_AllowPrinting, RibbonPermissionsTab.RibbonItem.TXITEM_AllowPrinting.ToString())
            Me.SetItemImage(Me.m_miEdit_Permissions_AllowCopy, RibbonPermissionsTab.RibbonItem.TXITEM_AllowCopy.ToString())
            Me.SetItemImage(Me.m_miEdit_Permissions_AllowEditingFormFields, RibbonPermissionsTab.RibbonItem.TXITEM_FillInFormFields.ToString())
            Me.SetItemImage(Me.m_miEdit_Permissions_ReadOnly, RibbonPermissionsTab.RibbonItem.TXITEM_ReadOnly.ToString())

            ' 'Editable Regions'
            Me.SetItemImage(Me.m_miEdit_EditableRegions, RibbonPermissionsTab.RibbonItem.TXITEM_HighlightEditableRegions.ToString())

            ' 'Protect Document'
            Me.SetItemImage(Me.m_miEdit_ProtectDocument, RibbonPermissionsTab.RibbonItem.TXITEM_EnforceProtection.ToString())

            ' 'Tracked Changes'
            Me.SetItemImage(Me.m_miEdit_TrackedChanges, RibbonProofingTab.RibbonItem.TXITEM_TrackChanges.ToString())
            Me.SetItemImage(Me.m_miEdit_TrackedChanges_TrackChanges, RibbonProofingTab.RibbonItem.TXITEM_TrackChanges.ToString())
            Me.SetItemImage(Me.m_miEdit_TrackedChanges_ReviewTrackedChanges, RibbonProofingTab.RibbonItem.TXITEM_TrackedChanges.ToString())

            ' 'Comments'
            Me.SetItemImage(Me.m_miEdit_Comments, RibbonProofingTab.RibbonItem.TXITEM_EditComment.ToString())
            Me.SetItemImage(Me.m_miEdit_Comments_AddComment, RibbonProofingTab.RibbonItem.TXITEM_AddComment.ToString())
            Me.SetItemImage(Me.m_miEdit_Comments_ReviewComments, RibbonProofingTab.RibbonItem.TXITEM_Comments_Sidebars.ToString())
        End Sub
    End Class
End Namespace
