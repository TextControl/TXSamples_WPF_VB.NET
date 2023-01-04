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
        Private m_dpdSidebarIsShown As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(Sidebar.IsShownProperty, GetType(Sidebar))
        Private m_dpdSidebarContentLayout As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(Sidebar.ContentLayoutProperty, GetType(Sidebar))


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
            m_dpdSidebarContentLayout.AddValueChanged(Me.m_sbSidebarLeft, New EventHandler(AddressOf SidebarLeft_ContentLayoutChanged))

            ' Right sidebar:
            m_dpdSidebarContentLayout.AddValueChanged(Me.m_sbSidebarRight, New EventHandler(AddressOf SidebarRight_ContentLayoutChanged))

            ' Bottom sidebar:
            m_dpdSidebarContentLayout.AddValueChanged(Me.m_sbSidebarBottom, New EventHandler(AddressOf SidebarBottom_ContentLayoutChanged))
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SidebarLeft_ContentLayoutChanged Handler
        ' Manages the layout of the left sidebar if its ContentLayout is set to TrackedChanges, DocumentSettings or 
        ' About.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SidebarLeft_ContentLayoutChanged(ByVal sender As Object, ByVal e As EventArgs)
            Select Case Me.m_sbSidebarLeft.ContentLayout
                Case Sidebar.SidebarContentLayout.TrackedChanges
                    Me.m_sbSidebarLeft.ShowPinButton = True
                    Me.m_rtbtnDocumentSettings.IsChecked = False
                    Me.m_rtbtnAbout.IsChecked = False
                Case Sidebar.SidebarContentLayout.DocumentSettings
                    Me.m_sbSidebarLeft.ShowPinButton = False
                    Me.m_sbSidebarLeft.IsPinned = True
                    Me.m_rtbtnDocumentSettings.IsChecked = True
                    Me.m_rtbtnAbout.IsChecked = False
                Case Sidebar.SidebarContentLayout.About
                    Me.m_sbSidebarLeft.ShowPinButton = False
                    Me.m_sbSidebarLeft.IsPinned = True
                    Me.m_rtbtnDocumentSettings.IsChecked = False
            End Select
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SidebarLeft_IsShownChanged Handler
        ' Toggles the document settings button if the left sidebar is show and its ContentLayout is set to 
        ' DocumentSettings. Otherwise the button is untoggled.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SidebarLeft_IsShownChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me.m_rtbtnDocumentSettings.IsChecked = Me.m_sbSidebarLeft.ContentLayout = Sidebar.SidebarContentLayout.DocumentSettings AndAlso Me.m_sbSidebarLeft.IsShown
            Me.m_rtbtnAbout.IsChecked = Me.m_sbSidebarLeft.ContentLayout = Sidebar.SidebarContentLayout.About AndAlso Me.m_sbSidebarLeft.IsShown
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SidebarRight_ContentLayoutChanged Handler
        ' Manages the layout of the right sidebar if its ContentLayout is set to ConditionalInstructions,
        ' FieldNavigator, Styles, Comments, Find or Replace.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SidebarRight_ContentLayoutChanged(ByVal sender As Object, ByVal e As EventArgs)

            Select Case Me.m_sbSidebarRight.ContentLayout
                Case Sidebar.SidebarContentLayout.ConditionalInstructions
                    Me.m_sbSidebarRight.ShowPinButton = True
                    Me.m_sbSidebarRight.DialogStyle = Sidebar.SidebarDialogStyle.Standard

                Case Sidebar.SidebarContentLayout.FieldNavigator, Sidebar.SidebarContentLayout.Styles
                    Me.m_sbSidebarRight.ShowPinButton = False
                    Me.m_sbSidebarRight.IsPinned = True
                    Me.m_sbSidebarRight.DialogStyle = Sidebar.SidebarDialogStyle.Standard
                Case Sidebar.SidebarContentLayout.Comments
                    Me.m_sbSidebarRight.ShowPinButton = True
                    Me.m_sbSidebarRight.DialogStyle = Sidebar.SidebarDialogStyle.StandardSizable
                Case Else
                    ' Find or Replace
                    Me.m_sbSidebarRight.ShowPinButton = True
                    Me.m_sbSidebarRight.DialogStyle = Sidebar.SidebarDialogStyle.Standard
            End Select
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SidebarBottom_ContentLayoutChanged Handler
        ' Manages the layout of the bottom sidebar if its ContentLayout is set to TrackedChanges, Comments, Find,
        ' Replace and GoTo.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SidebarBottom_ContentLayoutChanged(ByVal sender As Object, ByVal e As EventArgs)

            Select Case Me.m_sbSidebarBottom.ContentLayout
                Case Sidebar.SidebarContentLayout.TrackedChanges, Sidebar.SidebarContentLayout.Comments
                    Me.m_sbSidebarBottom.ShowTitle = True
                    Me.m_sbSidebarBottom.DialogStyle = Sidebar.SidebarDialogStyle.StandardSizable
                Case Else
                    ' Find, Replace and GoTo
                    Me.m_sbSidebarBottom.ShowTitle = False
                    Me.m_sbSidebarBottom.DialogStyle = Sidebar.SidebarDialogStyle.Standard
            End Select

        End Sub
    End Class
End Namespace
