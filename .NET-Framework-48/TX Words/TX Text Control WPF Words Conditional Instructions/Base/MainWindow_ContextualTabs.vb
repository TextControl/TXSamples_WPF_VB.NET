'-----------------------------------------------------------------------------------------------------------
' MainWindow_ContextualTabs.vb File
'
' Description:
'     Handles showing/hiding contextual tabs.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.DataVisualization

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetContextualTabsBehavior Method
        ' Sets the header of the contextual tabs and adds all necessary handlers to the TextControl.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetContextualTabsBehavior()
            ' Frame Tools:
            Me.m_ctgFrameTools.Header = My.Resources.ContextualTabGroup_FrameTools ' Set Frame Tools header

            ' Connect all necessary handlers to act as a Frame Tools group
            AddHandler Me.m_txTextControl.FrameSelected, AddressOf Me.TextControl_FrameSelected
            AddHandler Me.m_txTextControl.FrameDeselected, AddressOf Me.TextControl_FrameDeselected
            AddHandler Me.m_txTextControl.DrawingActivated, AddressOf Me.TextControl_DrawingActivated
            AddHandler Me.m_txTextControl.DrawingDeactivated, AddressOf Me.TextControl_DrawingDeactivated

            ' Table Tools:
            Me.m_ctgTableTools.Header = My.Resources.ContextualTabGroup_TableTools ' Set Table Tools header

            ' Connect all necessary handlers to act as a Table Tools group
            AddHandler Me.m_txTextControl.InputPositionChanged, AddressOf Me.TextControl_InputPositionChanged
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_InputPositionChanged Handler
        ' Checks whether the input position is in a table and makes the table layout tab visible.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_InputPositionChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me.m_ctgTableTools.Visibility = If(Me.m_txTextControl.Tables.GetItem() IsNot Nothing, Windows.Visibility.Visible, Windows.Visibility.Collapsed)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_FrameSelected Handler
        ' A frame has been selected. In this case make the frame layout tab visible.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_FrameSelected(ByVal sender As Object, ByVal e As FrameEventArgs)
            ' Show the Frame Tools group
            Me.m_ctgFrameTools.Visibility = Windows.Visibility.Visible
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_FrameDeselected Handler
        ' Makes the frame layout tab invisible. When a new frame is selected, the FrameSelected event of the
        ' new frame occurs before the FrameDeselected event of the old frame. Therefore it must be checked
        ' whether a new frame is selected. When a drawing is activated, the tab must also remain visible.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_FrameDeselected(ByVal sender As Object, ByVal e As FrameEventArgs)
            If Me.m_txTextControl.Frames.GetItem() Is Nothing AndAlso Me.m_txTextControl.Drawings.GetActivatedItem() Is Nothing Then
                ' If no frame is selected and no drawing is activated, hide the Frame Tools group
                Me.m_ctgFrameTools.Visibility = Windows.Visibility.Collapsed
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_DrawingActivated Handler
        ' When a drawing is activated, the contained shapes can also be formatted with the RibbonFrameLayoutTab.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_DrawingActivated(ByVal sender As Object, ByVal e As DrawingEventArgs)
            ' Show the Frame Tools group if the drawing is activated.
            Me.m_ctgFrameTools.Visibility = Windows.Visibility.Visible
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_DrawingDeactivated Handler
        ' Makes the frame layout tab invisible. When a frame is selected or another drawing is activated
        ' the tab must remain visible.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_DrawingDeactivated(ByVal sender As Object, ByVal e As DrawingEventArgs)
            If Me.m_txTextControl.Frames.GetItem() Is Nothing AndAlso Me.m_txTextControl.Drawings.GetActivatedItem() Is Nothing Then
                ' Hide the Frame Tools group if the drawing is deactivated and no other frame is selected.
                Me.m_ctgFrameTools.Visibility = Windows.Visibility.Collapsed
            End If
        End Sub
    End Class
End Namespace
