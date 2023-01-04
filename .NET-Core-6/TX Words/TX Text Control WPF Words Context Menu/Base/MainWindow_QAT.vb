'-----------------------------------------------------------------------------------------------------------
' MainWindow_QAT.vb File
'
' Description:
'      Creates an undo and a redo button and adds these button plus references to the [current user], Save,  
'      Open, New and Print buttons to the Quick Access Toolbar.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Windows.Controls.Ribbon
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetQATItemsDesign Method
        ' Creates an undo and a redo button and adds these button plus references to the [current user], Save, Open, 
        ' New and Print buttons to the Quick Access Toolbar.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetQATItemsDesign()
            Me.SetQATItemDesign(ResourceProvider.FileMenuItem.TXITEM_Save.ToString(), Me.m_rbtnSaveQAT, "1")
            Me.SetQATItemDesign(ResourceProvider.FileMenuItem.TXITEM_Open.ToString(), Me.m_rbtnOpenQAT, "2")
            Me.SetQATItemDesign(ResourceProvider.FileMenuItem.TXITEM_New.ToString(), Me.m_rbtnNewQAT, "3")
            Me.SetQATItemDesign(ResourceProvider.GeneralItem.TXITEM_Undo.ToString(), Me.m_rbtnUndoQAT, "4")
            Me.SetQATItemDesign(ResourceProvider.GeneralItem.TXITEM_Redo.ToString(), Me.m_rbtnRedoQAT, "5")
            Me.SetQATItemDesign(ResourceProvider.FileMenuItem.TXITEM_Print.ToString(), Me.m_rbtnPrintQAT, "6")
        End Sub

        '-----------------------------------------------------------------------------------------------------
        ' SetQATItemDesign Method
        ' Sets the icons, text, key tip and tool tip for a specific QAT RibbonButton.
        '
        ' Parameters:
        '      resourceID:     The id that is used to identify the corresponding texts and icons.
        '      menuItem:	   The ribbon menu item to update.
        '      keyTip:         The key tip to set.
        '-----------------------------------------------------------------------------------------------------
        Private Sub SetQATItemDesign(ByVal resourceID As String, ByVal menuItem As RibbonButton, ByVal keyTip As String)
            menuItem.Name = resourceID
            menuItem.SmallImageSource = ResourceProvider.GetSmallIcon(resourceID, Me)
            menuItem.LargeImageSource = ResourceProvider.GetLargeIcon(resourceID, Me)
            menuItem.KeyTip = keyTip

            menuItem.Label = ResourceProvider.GetText(resourceID)
            menuItem.ToolTipTitle = ResourceProvider.GetToolTipTitle(resourceID)
            menuItem.ToolTipDescription = ResourceProvider.GetToolTipDescription(resourceID)
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Undo_Click Handler
        ' Invokes the TextControl Undo method to undo the last action.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Undo_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Undo()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Undo_ToolTip_Opening Handler
        ' Sets the tool tip of the Undo button when the tool tip is opening. The tool tip shows the undo action
        ' that is performed when the Undo button is clicked.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Undo_ToolTip_Opening(ByVal sender As Object, ByVal e As ToolTipEventArgs)
            Dim strUndoActionName As String = Me.m_txTextControl.UndoActionName
            Me.m_rbtnUndoQAT.ToolTipDescription = If(Not String.IsNullOrEmpty(strUndoActionName), strUndoActionName, ResourceProvider.GetToolTipDescription(ResourceProvider.GeneralItem.TXITEM_Undo.ToString()))
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Redo_Click Handler
        ' Invokes the TextControl Redo method to redo the last action.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Redo_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Redo()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Redo_ToolTip_Opening Handler
        ' Sets the tool tip of the Redo button when the tool tip is opening. The tool tip shows the redo action
        ' that is performed when the Redo button is clicked.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Redo_ToolTip_Opening(ByVal sender As Object, ByVal e As ToolTipEventArgs)
            Dim strRedoActionName As String = Me.m_txTextControl.RedoActionName
            Me.m_rbtnRedoQAT.ToolTipDescription = If(Not String.IsNullOrEmpty(strRedoActionName), strRedoActionName, ResourceProvider.GetToolTipDescription(ResourceProvider.GeneralItem.TXITEM_Redo.ToString()))
        End Sub
    End Class
End Namespace
