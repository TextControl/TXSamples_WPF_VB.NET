'-----------------------------------------------------------------------------------------------------------
' MainWindow_Shortcuts.vb File
'
' Description:
'     Handles shortcuts.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_KeyDown Handler
        ' Implement Shortcuts for the TextControl.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            Select Case e.Key
                Case Key.Insert   ' Toggle insertion mode					
                    If Keyboard.Modifiers = ModifierKeys.None Then
                        ToggleInsertionMode()
                    End If

                Case Key.A        ' Ctrl + A: Select all
                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        Me.m_txTextControl.SelectAll()
                    End If

                Case Key.S        ' Ctrl + S: Save
                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        Save(m_strActiveDocumentPath)
                    End If

                Case Key.O        ' Ctrl + O: Open
                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        Open()
                    End If

                Case Key.F        ' Ctrl + F: Search
                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        Me.m_txTextControl.Find()
                    End If

                Case Key.P        ' Ctrl + P: Print
                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        If Me.m_txTextControl.CanPrint Then
                            Me.m_txTextControl.Print(m_strActiveDocumentName, True)
                        End If
                    End If
            End Select
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetShortCutBehavior Method
        ' Adds all necessary handlers to implement short cut behavior.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetShortCutBehavior()
            AddHandler Me.m_txTextControl.KeyDown, AddressOf Me.TextControl_KeyDown ' Handles shortcuts
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' ToggleInsertionMode Method
        ' Switch the TextControl's insertion mode between overwrite and insert.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub ToggleInsertionMode()
            Me.m_txTextControl.InsertionMode = If(Me.m_txTextControl.InsertionMode = InsertionMode.Insert, InsertionMode.Overwrite, InsertionMode.Insert)
        End Sub
    End Class
End Namespace
