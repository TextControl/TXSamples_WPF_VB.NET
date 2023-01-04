'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'		Sample project that is related to the 'Howto: Manipulate the MiniToolbar' article inside
'		the 'Windows Presentation Foundation User's Guide'.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Windows.Controls.Ribbon
Class MainWindow
    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_Loaded Handler
    ' Load the sample document.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.m_txTextControl.Load("Files\Sample.tx", TXTextControl.StreamType.InternalFormat)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_TextMiniToolbarInitialized Handler
    ' Modify the basic structure of the TextMiniToolbar.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_TextMiniToolbarInitalized(ByVal sender As Object, ByVal e As TXTextControl.WPF.MiniToolbarInitializedEventArgs)
        ' Ensure that the TextMiniToolbar's table layout group won't be displayed if the input position is inside a table.
        e.MiniToolbar.Container.Children.Remove(TryCast(e.MiniToolbar.Container.FindName(TXTextControl.WPF.TextMiniToolbar.RibbonItem.TXITEM_TableLayoutGroup.ToString()), UIElement))

        ' Add a ribbon group separator.
        e.MiniToolbar.Container.Children.Add(CreateRibbonGroupSeperator(3))

        ' Create and add a ribbon group to the TextMiniToolbar that provides an "Edit Hyperlink" button.
        e.MiniToolbar.Container.ColumnDefinitions.Add(New ColumnDefinition() With {
                .Width = GridLength.Auto
            })
        e.MiniToolbar.Container.Children.Add(CreateEditHyperlinkGroup(4))
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_MiniToolbarOpening Handler
    ' Update the TextMiniToolbar's content visibility.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_MiniToolbarOpening(ByVal sender As Object, ByVal e As TXTextControl.WPF.MiniToolbarOpeningEventArgs)
        ' Check whether the opening mini tool bar is of type TextMiniToolbar
        If TypeOf e.MiniToolbar Is TXTextControl.WPF.TextMiniToolbar Then
            e.MiniToolbar.Container.Children(2).Visibility = Visibility.Visible ' Ensure that the TextMiniToolbar's Styles group is always shown (even if the input position is inside a table)

            e.MiniToolbar.Container.Children(3).Visibility = CSharpImpl.Assign(e.MiniToolbar.Container.Children(4).Visibility, If((e.MiniToolbarContext And TXTextControl.ContextMenuLocation.TextField) = TXTextControl.ContextMenuLocation.TextField AndAlso Me.m_txTextControl.HypertextLinks.GetItem() IsNot Nothing, Visibility.Visible, Visibility.Collapsed)) ' Ensure that the ribbon group's separator and...
            ' ... the "Edit Hyperlink" group are displayed if...
            ' ... the current context is TextField and ...
            ' ... the text field is of type TXTextControl.HypertextLink.
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' EditHyperlink_Click Handler
    ' Opens the TXTextControl HyperlinkDialog.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub EditHyperlink_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim dlgHyperlinkDialog As TXTextControl.WPF.HyperlinkDialog = New TXTextControl.WPF.HyperlinkDialog(Me.m_txTextControl)
        dlgHyperlinkDialog.Owner = Me
        dlgHyperlinkDialog.ShowDialog()
    End Sub


    '-----------------------------------------------------------------------------------------------------------
    ' M E T H O D S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' CreateRibbonGroupSeperator Method
    ' Creates a ribbon group separator.
    '-----------------------------------------------------------------------------------------------------------
    Protected Function CreateRibbonGroupSeperator(ByVal column As Integer) As RibbonSeparator
        ' Create ribbon separator
        Dim rsRibbonGroupSeperator As RibbonSeparator = New RibbonSeparator()
        rsRibbonGroupSeperator.LayoutTransform = New RotateTransform(90)
        rsRibbonGroupSeperator.Margin = New Thickness(2)
        rsRibbonGroupSeperator.BorderThickness = New Thickness(5)

        ' Set row and column
        Grid.SetRow(rsRibbonGroupSeperator, 0)
        Grid.SetRowSpan(rsRibbonGroupSeperator, 3)
        Grid.SetColumn(rsRibbonGroupSeperator, column)

        Return rsRibbonGroupSeperator
    End Function

    '-----------------------------------------------------------------------------------------------------------
    ' CreateEditHyperlinkGroup Method
    ' Create a Grid that contains an Edit Hyperlink button.
    '
    ' Parameters:
    '		column		The column where to add the Grid.
    '
    ' Returns:			The created Grid.	
    '-----------------------------------------------------------------------------------------------------------
    Private Function CreateEditHyperlinkGroup(ByVal column As Integer) As Grid
        ' Create a ribbon group (represented by a Grid) that contains... 
        Dim rgEditHyperlinkGroup As Grid = New Grid()
        Grid.SetRow(rgEditHyperlinkGroup, 0)
        Grid.SetRowSpan(rgEditHyperlinkGroup, 3)
        Grid.SetColumn(rgEditHyperlinkGroup, column)

        rgEditHyperlinkGroup.ColumnDefinitions.Add(New ColumnDefinition() With {
                .Width = GridLength.Auto
            })

        ' ... a button to open the TextControl Edit HyperlinkDialog
        Dim rbtnEditHyperlinkButton As RibbonButton = New RibbonButton() With {
                .Label = "Edit Hyperlink"
            }
        Grid.SetRow(rbtnEditHyperlinkButton, 0)
        Grid.SetColumn(rbtnEditHyperlinkButton, 0)

        rbtnEditHyperlinkButton.LargeImageSource = TXTextControl.WPF.ResourceProvider.GetLargeIcon(TXTextControl.WPF.RibbonInsertTab.RibbonItem.TXITEM_InsertHyperlink.ToString(), Me)
        AddHandler rbtnEditHyperlinkButton.Click, AddressOf EditHyperlink_Click

        ' Add the edit hyperlink button to group.
        rgEditHyperlinkGroup.Children.Add(rbtnEditHyperlinkButton)

        Return rgEditHyperlinkGroup
    End Function


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