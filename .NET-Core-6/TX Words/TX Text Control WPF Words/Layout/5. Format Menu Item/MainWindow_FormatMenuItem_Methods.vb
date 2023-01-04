'-----------------------------------------------------------------------------------------------------------
' MainWindow_FormatMenuItem_Methods.vb File
'
' Description: Provides supporting methods to implement the desired behavior of some 'Format' menu items.
'     
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.WPF.Drawing

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' 'Bullets and Numbering' Item
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetNumberedList Method
        ' Sets whether there is a numbered or structured list at the current input position. If such a list is set,
        ' the specified list format is set.
        '
        ' Parameters:
        '      isChecked:	Determines whether a numbered or structured list is set at the current input position.
        '		listFormat:	Specifies the list format to set.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetNumberedList(ByVal isChecked As Boolean, ByVal listFormat As NumberFormat)
            ' Check whether a structured or a numbered list is handled.
            If Me.m_miFormat_BulletsAndNumbering_AsStructuredList.IsChecked Then
                ' Set or remove the structured list.
                If CSharpImpl.Assign(Me.m_txTextControl.InputFormat.StructuredList, isChecked).Value Then
                    ' If a structured list is set, determine its list format.
                    Me.m_txTextControl.InputFormat.StructuredListFormat = listFormat
                End If
            Else
                ' Set or remove the numbered list.
                If CSharpImpl.Assign(Me.m_txTextControl.InputFormat.NumberedList, isChecked).Value Then
                    ' If a numbered list is set, determine its list format.
                    Me.m_txTextControl.InputFormat.NumberedListFormat = listFormat
                End If
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' 'Shape...' Item
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' EnableShapeItem Method
        ' Returns a value indicating whether the specified FrameBase element is a DrawingFrame or there is an
        ' activated DrawingFrame where at least one shape is selected.
        '
        ' Parameters:
        '      frame:		The FrameBase element to check.
        '
        ' Return Value:	True, if the specified FrameBase element is a DrawingFrame or there is an activated 
        '					DrawingFrame where at least one shape is selected. Otherwise false.
        '-----------------------------------------------------------------------------------------------------------
        Private Function EnableShapeItem(ByVal frame As FrameBase) As Boolean
            ' Check whether the specified FrameBase element is a DrawingFrame.
            Dim bEnableDrawingItem = TypeOf frame Is DataVisualization.DrawingFrame

            If frame Is Nothing Then
                ' Check whether there is a DrawingFrame that is currently activated.
                Dim dfActivatedDrawingFrame As DataVisualization.DrawingFrame = Me.m_txTextControl.Drawings.GetActivatedItem()

                If dfActivatedDrawingFrame IsNot Nothing Then
                    ' Check whether the activated DrawingFrame contains at least on selected shape.
                    Dim txdDrawingControl As TXDrawingControl = TryCast(dfActivatedDrawingFrame.Drawing, TXDrawingControl)
                    bEnableDrawingItem = txdDrawingControl.Selection.Shapes.Length > 0
                End If
            End If

            Return bEnableDrawingItem
        End Function
    End Class
End Namespace
