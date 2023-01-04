'-----------------------------------------------------------------------------------------------------------
' MainWindow_ContextMenu.vb File
'
' Description:
'      Adds an item to the TextControl context menu when a frame is selected. Clicking that item opens
'      a dialog to edit the name of the selected frame.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports TXTextControl.DataVisualization
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_TextContextMenuOpening Handler
        ' Customize the context menu by adding additional items if a frame is selected.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_TextContextMenuOpening(ByVal sender As Object, ByVal e As TextContextMenuEventArgs)
            If (e.ContextMenuLocation And ContextMenuLocation.SelectedFrame) <> 0 Then
                ' A frame is selected
                AddFrameContextMenuItems(e.TextContextMenu)
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SetFrameName_Click Handler
        ' Show a dialog on click for setting the FrameBase object's name.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetFrameName_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Get the current frame base object
            Dim fbFrame As FrameBase
            If (CSharpImpl.Assign(fbFrame, Me.m_txTextControl.Frames.GetItem())) Is Nothing Then
                Return
            End If

            ' Initialize a UserPromptDialog to edit the frame base name 
            Dim strName = fbFrame.Name
            Dim dlgFrameNameDialog As FrameNameDialog = New FrameNameDialog(strName)
            Dim bDialogResult As Boolean? = dlgFrameNameDialog.ShowDialog()
            ' Open the dialog.
            If bDialogResult.HasValue AndAlso bDialogResult.Value Then
                ' Set the new name.
                fbFrame.Name = dlgFrameNameDialog.FrameName

                ' Update the ribbon frame layout's Name text box.           
                Dim tbxObjectName As TextBox = TryCast(Me.m_rtRibbonFrameLayoutTab.FindName(RibbonFrameLayoutTab.RibbonItem.TXITEM_ObjectName_Textbox.ToString()), TextBox)
                tbxObjectName.Text = fbFrame.Name
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetContextMenuBehavior Method
        ' Adds all necessary handlers to show a customized context menu.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetContextMenuBehavior()
            AddHandler Me.m_txTextControl.TextContextMenuOpening, AddressOf Me.TextControl_TextContextMenuOpening ' Adds context specific items to the TextControl 'Text Context Menu' 
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' AddFrameContextMenuItems Method
        ' Adds the new context menu items, to ContextMenuStrip, for setting the framebase's name.
        '
        ' Parameters:
        '		contextMenu:    The context menu where to add the new items.	
        '-----------------------------------------------------------------------------------------------------------
        Private Sub AddFrameContextMenuItems(ByVal contextMenu As ContextMenu)
            Dim frame As FrameBase = Me.m_txTextControl.Frames.GetItem()
            If frame IsNot Nothing AndAlso Not (TypeOf frame Is ChartFrame) AndAlso Me.m_txTextControl.CanEdit Then
                contextMenu.Items.Add(New Separator()) ' Add separator.
                Dim miFrameName As MenuItem = New MenuItem() With {
                    .Header = My.Resources.ContextMenu_FameName,  ' Get item text.
                    .Icon = New Windows.Controls.Image() With {
                        .Source = ResourceProvider.GetSmallIcon(RibbonFrameLayoutTab.RibbonItem.TXITEM_ObjectName.ToString(), Me)
                    }  ' Get item icon.
                }
                AddHandler miFrameName.Click, AddressOf SetFrameName_Click  ' Add Click event to open a dialog to edit the frame base name.
                contextMenu.Items.Add(miFrameName)
            End If
        End Sub
    End Class
End Namespace
