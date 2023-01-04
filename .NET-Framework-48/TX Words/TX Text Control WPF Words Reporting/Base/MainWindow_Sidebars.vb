'-----------------------------------------------------------------------------------------------------------
' MainWindow_Sidebars.vb File
'
' Description:
'     Manages the layout of the sidebars when the content layout changed.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Imports System.ComponentModel
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private m_dpdSidebarContentLayout As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(Sidebar.ContentLayoutProperty, GetType(Sidebar))
        Private m_dpdSidebarIsShown As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(Sidebar.IsShownProperty, GetType(Sidebar))


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetSidebarBehavior Method
        ' Connects the sidebars with the corresponding property value changed handlers.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetSidebarBehavior()
            ' Left sidebar:
            m_dpdSidebarIsShown.AddValueChanged(Me.m_sbSidebarLeft, New EventHandler(AddressOf SidebarLeft_IsShownChanged))

            ' Right sidebar:
            m_dpdSidebarContentLayout.AddValueChanged(Me.m_sbSidebarRight, New EventHandler(AddressOf SidebarRight_ContentLayoutChanged))

            ' Bottom sidebar:
            m_dpdSidebarContentLayout.AddValueChanged(Me.m_sbSidebarBottom, New EventHandler(AddressOf SidebarBottom_ContentLayoutChanged))
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '----------------------------------------------------------------------------------------------------------
        ' SidebarLeft_IsShownChanged Handler
        ' Toggles the 'About' button if the left sidebar is shown and its ContentLayout is set to 
        ' SidebarContentLayout.About.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SidebarLeft_IsShownChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me.m_rtbtnAbout.IsChecked = Me.m_sbSidebarLeft.ContentLayout = Sidebar.SidebarContentLayout.About AndAlso Me.m_sbSidebarLeft.IsShown
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SidebarRight_ContentLayoutChanged Handler
        ' Manages the layout of the right sidebar if its ContentLayout is set to FieldNavigator, Styles, Find or 
        ' Replace.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SidebarRight_ContentLayoutChanged(ByVal sender As Object, ByVal e As EventArgs)

            Select Case Me.m_sbSidebarRight.ContentLayout
                Case Sidebar.SidebarContentLayout.FieldNavigator, Sidebar.SidebarContentLayout.Styles
                    Me.m_sbSidebarRight.ShowPinButton = False
                    Me.m_sbSidebarRight.IsPinned = True
                    Me.m_sbSidebarRight.DialogStyle = Sidebar.SidebarDialogStyle.Standard
                Case Else
                    ' Find or Replace
                    Me.m_sbSidebarRight.ShowPinButton = True
                    Me.m_sbSidebarRight.DialogStyle = Sidebar.SidebarDialogStyle.Standard
            End Select
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SidebarBottom_ContentLayoutChanged Handler
        ' Manages the layout of the bottom sidebar if its ContentLayout is set to Find, Replace and GoTo.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SidebarBottom_ContentLayoutChanged(ByVal sender As Object, ByVal e As EventArgs)

            Select Case Me.m_sbSidebarBottom.ContentLayout
                Case Else
                    ' Find, Replace and GoTo
                    Me.m_sbSidebarBottom.ShowTitle = False
                    Me.m_sbSidebarBottom.DialogStyle = Sidebar.SidebarDialogStyle.Standard
            End Select

        End Sub
    End Class
End Namespace
