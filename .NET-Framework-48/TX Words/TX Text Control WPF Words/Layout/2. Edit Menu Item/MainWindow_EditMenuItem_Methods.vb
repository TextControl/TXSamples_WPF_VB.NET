'-----------------------------------------------------------------------------------------------------------
' MainWindow_EditMenuItem_Methods.vb File
'
' Description: Provides supporting methods to implement the desired behavior of some 'Edit' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' 'Undo' Item
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateUndoText Method
        ' Sets the Undo item text by combining the default 'Undo' text with the corresponding undo action.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateUndoText()
            ' Update Undo item text.
            Me.m_miEdit_Undo.Header = My.Resources.Item_Edit_Undo_Text

            ' Add undo action text that is performed when the Undo item is clicked.
            Dim strUndoActionName As String = Me.m_txTextControl.UndoActionName

            If Not String.IsNullOrEmpty(strUndoActionName) Then
                Me.m_miEdit_Undo.Header += " " & strUndoActionName
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' 'Redo' Item
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' UpdateRedoText Method
        ' Sets the Redo item text by combining the default 'Redo' text with the corresponding redo action.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub UpdateRedoText()
            ' Update Redo item text.
            Me.m_miEdit_Redo.Header = My.Resources.Item_Edit_Redo_Text

            ' Add redo action text that is performed when the Redo item is clicked.
            Dim strRedoActionName As String = Me.m_txTextControl.RedoActionName

            If Not String.IsNullOrEmpty(strRedoActionName) Then
                Me.m_miEdit_Redo.Header += " " & strRedoActionName
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' 'Editable Regions' Items
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' HandleAddEditableRegionFailure Method
        ' Handles the returned AddResult when adding a new editable region to the document. If the result is not
        ' AddResult.Successful, a corresponding message box with a detailed description of the failure is shown.
        '
        ' Parameters:
        '      addResult:		The EditableRegionCollection.AddResult value to check.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub HandleAddEditableRegionFailure(ByVal addResult As EditableRegionCollection.AddResult)
            If addResult <> EditableRegionCollection.AddResult.Successful Then
                Dim strMessageBoxCaption As String = Nothing
                Dim strMessageBoxText As String = Nothing

                Select Case addResult
                    Case EditableRegionCollection.AddResult.Error
                        strMessageBoxCaption = My.Resources.MessageBox_AddEditableRegion_Error_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddEditableRegion_Error_Text
                    Case EditableRegionCollection.AddResult.NoSelection
                        strMessageBoxCaption = My.Resources.MessageBox_AddEditableRegion_NoSelection_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddEditableRegion_NoSelection_Text
                    Case EditableRegionCollection.AddResult.PositionInvalid
                        strMessageBoxCaption = My.Resources.MessageBox_AddEditableRegion_PositionInvalid_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddEditableRegion_PositionInvalid_Text
                    Case EditableRegionCollection.AddResult.SelectionTooComplex
                        strMessageBoxCaption = My.Resources.MessageBox_AddEditableRegion_SelectionTooComplex_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddEditableRegion_SelectionTooComplex_Text
                End Select

                If Not Equals(strMessageBoxText, Nothing) Then
                    MessageBox.Show(Me, strMessageBoxText, strMessageBoxCaption, MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' RemoveUser Method
        ' Removes the editable region that is related to the specified user name from the current text position.
        '
        ' Parameters:
        '      userName:		The name of the user that is related to the editable region that should be removed. 
        '						If the parameter is an empty string, the editable region is related to 'Everyone'.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub RemoveUser(ByVal userName As String)
            Dim colEditableRegions As EditableRegionCollection = Me.m_txTextControl.EditableRegions
            Dim rrrCurrentEditableRegions As EditableRegion() = colEditableRegions.GetItems()
            Dim erEditableRegionToRemove = GetEditableRegion(userName, rrrCurrentEditableRegions)
            colEditableRegions.Remove(erEditableRegionToRemove, Me.m_txTextControl.Selection.Length > 0)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' GetEditableRegion Method
        ' Searches in an array of EditableRegion objects for an editable region that corresponds to the name of a 
        ' specific user.
        '
        ' Parameters:
        '      userName:			The name of the specified user.
        '      editableRegions:	An array of EditableRegion objects that are checked.
        '
        ' Return Value:			If found, the editable Region that corresponds to the name of the specified user. 
        '							Otherwise null.
        '-----------------------------------------------------------------------------------------------------------
        Private Function GetEditableRegion(ByVal userName As String, ByVal editableRegions As EditableRegion()) As EditableRegion
            For i = 0 To editableRegions.Length - 1

                If Equals(editableRegions(i).UserName, userName) Then
                    Return editableRegions(i)
                End If
            Next

            Return Nothing
        End Function


        '-----------------------------------------------------------------------------------------------------------
        ' 'Comments' Items
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' HandleAddCommentFailure Method
        ' Handles the returned AddResult when adding a new comment to the document. If the result is not
        ' AddResult.Successful, a corresponding message box with a detailed description of the failure is shown.
        '
        ' Parameters:
        '      addResult:		The CommentCollection.AddResult value to check.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub HandleAddCommentFailure(ByVal addResult As CommentCollection.AddResult)
            If addResult <> CommentCollection.AddResult.Successful Then
                Dim strMessageBoxCaption As String = Nothing
                Dim strMessageBoxText As String = Nothing

                Select Case addResult
                    Case CommentCollection.AddResult.Error
                        strMessageBoxCaption = My.Resources.MessageBox_AddComment_Error_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddComment_Error_Text
                    Case CommentCollection.AddResult.NoSelection
                        strMessageBoxCaption = My.Resources.MessageBox_AddComment_NoSelection_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddComment_NoSelection_Text
                    Case CommentCollection.AddResult.PositionInvalid
                        strMessageBoxCaption = My.Resources.MessageBox_AddComment_PositionInvalid_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddComment_PositionInvalid_Text
                    Case CommentCollection.AddResult.SelectionTooComplex
                        strMessageBoxCaption = My.Resources.MessageBox_AddComment_SelectionTooComplex_Caption
                        strMessageBoxText = My.Resources.MessageBox_AddComment_SelectionTooComplex_Text
                End Select

                If Not Equals(strMessageBoxText, Nothing) Then
                    MessageBox.Show(Me, strMessageBoxText, strMessageBoxCaption, MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            End If
        End Sub
    End Class
End Namespace
