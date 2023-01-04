'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'     Implements TX Text Control Words.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.ComponentModel
Imports System.Reflection
Imports System.Windows.Controls.Ribbon
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits RibbonWindow
        ' ----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_strFilesDirectory As String = "Files\"


        '-----------------------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' MainWindow Constructor
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New()
            ' Add an unhandled exception handler
            Dim currentDomain As AppDomain = AppDomain.CurrentDomain
            AddHandler currentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException

            Me.InitializeComponent()

            ' Set some texts
            Me.Title = My.Resources.MainWindow_Caption_Product
            Me.m_rtRibbonTableLayoutTab.ContextualTabGroupHeader = My.Resources.ContextualTabGroup_TableTools
            Me.m_rtRibbonFormulaTab.ContextualTabGroupHeader = My.Resources.ContextualTabGroup_TableTools
            Me.m_rtRibbonFrameLayoutTab.ContextualTabGroupHeader = My.Resources.ContextualTabGroup_FrameTools
            Me.m_ctgTableTools.Header = My.Resources.ContextualTabGroup_TableTools
            Me.m_ctgFrameTools.Header = My.Resources.ContextualTabGroup_FrameTools

            ' Get and set saved application settings.
            LoadRecentFiles()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------
        ' SetAppMenuDesign Method
        ' Updates the design of the menu items.
        '-----------------------------------------------------------------------------------------------------
        Private Sub SetAppMenuDesign()
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_New.ToString(), Me.m_rmiNew)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Open.ToString(), Me.m_rmiOpen)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Save.ToString(), Me.m_rmiSave)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_SaveAs.ToString(), Me.m_rmiSaveAs)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Print.ToString(), Me.m_rsmiPrint)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Print.ToString(), Me.m_rbtnPrint)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_PrintQuick.ToString(), Me.m_rbtnPrintQuick)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_About.ToString(), Me.m_rtbtnAbout, My.Resources.AboutButton_ToolTip_Description)

            ' Set Recent Files overview label text
            Me.m_rgcRecentFiles.Header = My.Resources.ApplicationMenu_ResentFiles
        End Sub

        '-----------------------------------------------------------------------------------------------------
        ' SetRibbonButtonDesign Method
        ' Sets the icons, text, key tip and tool tip for a specific RibbonButton.
        '
        ' Parameters:
        '      resourceID:     The id that is used to identify the corresponding texts and icons.
        '      menuItem:	   The ribbon menu item to update.
        '      args:           Optional strings that are used to format the displayed texts.
        '-----------------------------------------------------------------------------------------------------
        Private Sub SetMenuItemDesign(ByVal resourceID As String, ByVal menuItem As RibbonMenuItem, ParamArray args As String())
            menuItem.Name = resourceID
            menuItem.ImageSource = ResourceProvider.GetLargeIcon(resourceID, Me)
            menuItem.KeyTip = ResourceProvider.GetKeyTip(resourceID)

            If args.Length > 0 Then
                menuItem.Header = String.Format(ResourceProvider.GetText(resourceID), args)
                menuItem.ToolTipTitle = String.Format(ResourceProvider.GetToolTipTitle(resourceID), args)
                menuItem.ToolTipDescription = String.Format(ResourceProvider.GetToolTipDescription(resourceID), args)
            Else
                menuItem.Header = ResourceProvider.GetText(resourceID)
                menuItem.ToolTipTitle = ResourceProvider.GetToolTipTitle(resourceID)
                menuItem.ToolTipDescription = ResourceProvider.GetToolTipDescription(resourceID)
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------
        ' SetAppMenuBehavior Method
        ' Connects all necessary handlers to the application menu items.
        '-----------------------------------------------------------------------------------------------------
        Private Sub SetAppMenuBehavior()
            ' Common
            AddHandler Me.m_txTextControl.Changed, AddressOf Me.TextControl_Changed ' Updates the internal 'is dirty document' flag

            ' New:
            AddHandler Me.m_rmiNew.Click, AddressOf Me.New_Click

            ' Open:
            AddHandler Me.m_rmiOpen.Click, AddressOf Me.Open_Click

            ' Save:
            AddHandler Me.m_rmiSave.Click, AddressOf Me.Save_Click

            ' Save As:
            AddHandler Me.m_rmiSaveAs.Click, AddressOf Me.SaveAs_Click

            ' Print:
            AddHandler Me.m_rsmiPrint.Click, AddressOf Me.Print_Click ' Print(Split Button)
            AddHandler Me.m_rbtnPrint.Click, AddressOf Me.Print_Click ' Print (Drop Down Button)
            AddHandler Me.m_rbtnPrintQuick.Click, AddressOf Me.PrintQuick_Click ' Print Quick (Drop Down Button)
            AddHandler Me.m_txTextControl.PropertyChanged, AddressOf Me.TextControl_PropertyChanged_Print ' Add handler for the print buttons Enable state handling
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------
        ' MainWindow_Loaded Handler 
        ' Sets the requested behavior for all added controls.
        '-----------------------------------------------------------------------------------------------------
        Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Main Window:
            UpdateMainWindowCaption() ' Set caption

            ' Application Menu items:
            SetAppMenuDesign()
            SetAppMenuBehavior()

            ' QAT:
            SetQATItemsDesign()

            ' Sidebars:
            SetSidebarBehavior()

            ' Mini Toolbar
            m_txTextControl.ShowMiniToolbar = MiniToolbarButton.LeftButton Or MiniToolbarButton.RightButton

            ' Contextual Tabs:
            SetContextualTabsBehavior()

            ' Drag & Drop
            SetDragAndDropBehavior()

            ' Toolbars:
            SetRulerBarsDesign()
            SetStatusBarDesign()

            ' About:
            Me.UpdateAboutSidebar()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_Loaded_MainWindow Handler 
        ' Sets the focus to the TextControl.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_Loaded_MainWindow(ByVal sender As Object, ByVal e As RoutedEventArgs)
            m_txTextControl.Focus()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' MainWindow_Closing Handler
        ' Saves the recent files to the My.Settings.Default.RecentFiles property.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As CancelEventArgs)
            ' Save the recent files to the My.Settings.Default.RecentFiles property
            SaveRecentFiles()
        End Sub


        '-----------------------------------------------------------------------------------------------------
        ' CurrentDomain_UnhandledException Handler
        ' Handles an exception by showing a message box that explains the reason for the exception.
        '-----------------------------------------------------------------------------------------------------
        Friend Shared Sub CurrentDomain_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
            Dim strProductName As String = CType(Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), GetType(AssemblyProductAttribute)), AssemblyProductAttribute).Product
            Dim ex As Exception = CType(e.ExceptionObject, Exception)

            ' TX Text Control Feature is not available
            If TypeOf ex Is LicenseLevelException Then
                MessageBox.Show(String.Format(My.Resources.MessageBox_Application_ThreadException_Text, ex.Message), strProductName, MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            ' All other exceptions
            MessageBox.Show(String.Format(My.Resources.MessageBox_Application_ThreadException_Text, ex.Message), strProductName, MessageBoxButton.OK, MessageBoxImage.Error)
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' C L A S S E S
        '-----------------------------------------------------------------------------------------------------------
        Private Class CSharpImpl
            Shared Function Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
